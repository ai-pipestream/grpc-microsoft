package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DownloadItemRequest;
import ai.pipestream.microsoft.v1.DownloadItemResponse;
import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.GetItemResponse;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListChildrenResponse;
import ai.pipestream.microsoft.v1.ListDrivesRequest;
import ai.pipestream.microsoft.v1.ListDrivesResponse;
import ai.pipestream.microsoft.v1.ListSitesRequest;
import ai.pipestream.microsoft.v1.ListSitesResponse;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.ByteString;
import com.google.protobuf.Message;
import io.grpc.Status;
import io.grpc.stub.StreamObserver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * {@code MicrosoftService} facade over Graph: Jackson mapper plus the
 * crawler. Outgoing messages are checked by {@link MicrosoftValidator}.
 */
public final class MicrosoftGrpcService extends MicrosoftServiceGrpc.MicrosoftServiceImplBase {

    /** Default cap for inlined file bytes: 25 MiB. */
    public static final long DEFAULT_ATTACHMENT_MAX_BYTES = 25L * 1024 * 1024;

    private static final MicrosoftValidator VALIDATOR = MicrosoftValidator.create();

    private final MicrosoftConnectorConfig config;
    private final GraphFiles files;
    private final MicrosoftMapper mapper;
    private final long attachmentMaxBytes;
    private final MicrosoftChangeSink downstream;

    /**
     * Creates a service with no downstream {@link MicrosoftChangeSink}.
     *
     * @param config connector config
     * @param files the authorized files API
     * @param attachmentMaxBytes inline file byte cap; must be positive
     */
    public MicrosoftGrpcService(MicrosoftConnectorConfig config, GraphFiles files,
            long attachmentMaxBytes) {
        this(config, files, attachmentMaxBytes, null);
    }

    /**
     * Creates a service. When {@code downstream} is non-null, {@code Sync} also
     * fans out to that sink.
     *
     * @param config connector config
     * @param files the authorized files API
     * @param attachmentMaxBytes inline file byte cap; must be positive
     * @param downstream optional extra sink for {@code Sync} emissions;
     *        {@code null} for none
     */
    public MicrosoftGrpcService(MicrosoftConnectorConfig config, GraphFiles files,
            long attachmentMaxBytes, MicrosoftChangeSink downstream) {
        this.config = Objects.requireNonNull(config, "config");
        this.files = Objects.requireNonNull(files, "files");
        this.mapper = new MicrosoftMapper();
        if (attachmentMaxBytes <= 0) {
            throw new IllegalArgumentException("attachmentMaxBytes must be positive");
        }
        this.attachmentMaxBytes = attachmentMaxBytes;
        this.downstream = downstream;
    }

    @Override
    public void getMe(GetMeRequest request, StreamObserver<GetMeResponse> observer) {
        try {
            GraphUser user = requireValid(mapper.toUser(files.me()));
            observer.onNext(GetMeResponse.newBuilder().setUser(user).build());
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void listSites(ListSitesRequest request, StreamObserver<ListSitesResponse> observer) {
        try {
            String query = request.getQuery().isBlank() ? "*" : request.getQuery();
            JsonNode page = files.searchSites(query);
            List<Site> sites = new ArrayList<>();
            int cap = request.getLimit();
            for (JsonNode node : page.path("value")) {
                if (cap > 0 && sites.size() >= cap) {
                    break;
                }
                sites.add(requireValid(mapper.toSite(node)));
            }
            observer.onNext(ListSitesResponse.newBuilder().addAllSites(sites).build());
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void listDrives(ListDrivesRequest request, StreamObserver<ListDrivesResponse> observer) {
        try {
            List<Drive> drives = new ArrayList<>();
            if (request.getSiteId().isBlank()) {
                drives.add(requireValid(mapper.toDrive(files.meDrive())));
            } else {
                files.drives(request.getSiteId()).path("value")
                        .forEach(node -> drives.add(requireValid(mapper.toDrive(node,
                                request.getSiteId()))));
            }
            observer.onNext(ListDrivesResponse.newBuilder().addAllDrives(drives).build());
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void listChildren(ListChildrenRequest request,
            StreamObserver<ListChildrenResponse> observer) {
        try {
            for (JsonNode node : files.childrenAll(request.getDriveId(), request.getFolderPath())) {
                observer.onNext(ListChildrenResponse.newBuilder()
                        .setItem(requireValid(mapper.toDriveItem(node, request.getDriveId())))
                        .build());
            }
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void getItem(GetItemRequest request, StreamObserver<GetItemResponse> observer) {
        try {
            DriveItem item = requireValid(mapper.toDriveItem(
                    files.item(request.getDriveId(), request.getItemId()), request.getDriveId()));
            observer.onNext(GetItemResponse.newBuilder().setItem(item).build());
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void downloadItem(DownloadItemRequest request,
            StreamObserver<DownloadItemResponse> observer) {
        try {
            DriveItem item = requireValid(mapper.toDriveItem(
                    files.item(request.getDriveId(), request.getItemId()), request.getDriveId()));
            if (request.getIncludeContent()) {
                item = requireValid(withContent(item));
            }
            observer.onNext(DownloadItemResponse.newBuilder().setItem(item).build());
            observer.onCompleted();
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    @Override
    public void sync(SyncRequest request, StreamObserver<SyncResponse> observer) {
        try {
            MicrosoftConnectorConfig effective = request.getDriveIdsList().isEmpty()
                    ? config
                    : new MicrosoftConnectorConfig(config.tenantId(), config.clientId(),
                            config.clientSecret(), config.siteId(), request.getDriveIdsList(),
                            request.getFolderPath().isBlank() ? config.folderPath()
                                    : request.getFolderPath(),
                            config.graphBaseUrl(), config.authority());
            Object lock = new Object();
            java.util.concurrent.atomic.AtomicReference<String> runId =
                    new java.util.concurrent.atomic.AtomicReference<>("");
            MicrosoftChangeSink observerSink = new MicrosoftChangeSink() {
                @Override
                public void emit(MicrosoftChange change) {
                    runId.compareAndSet("", change.getCursor());
                    VALIDATOR.requireValid(change);
                    synchronized (lock) {
                        observer.onNext(SyncResponse.newBuilder().setChange(change).build());
                    }
                }

                @Override
                public void snapshot(MicrosoftSnapshot snapshot) {
                    VALIDATOR.requireValid(snapshot);
                    synchronized (lock) {
                        observer.onNext(SyncResponse.newBuilder().setSnapshot(snapshot).build());
                    }
                }
            };
            MicrosoftChangeSink sink = downstream == null ? observerSink
                    : new CompositeMicrosoftChangeSink(java.util.List.of(observerSink, downstream));
            MicrosoftCrawler crawler = new MicrosoftCrawler(effective, files, sink,
                    attachmentMaxBytes, request.getIncludeContent());
            crawler.crawl();
            sink.completeRun(runId.get());
            synchronized (lock) {
                observer.onNext(SyncResponse.newBuilder()
                        .setResumeCursor(java.time.Instant.now().toString()).build());
                observer.onCompleted();
            }
        } catch (Throwable t) {
            fail(observer, t);
        }
    }

    private DriveItem withContent(DriveItem item) throws IOException, InterruptedException {
        if (item.getFolder()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " is a folder; no binary")
                    .asRuntimeException();
        }
        if (item.getSize() > attachmentMaxBytes) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " is " + item.getSize()
                            + " bytes, above the " + attachmentMaxBytes + " byte inline cap")
                    .asRuntimeException();
        }
        byte[] bytes = files.download(item.getDriveId(), item.getId());
        if (bytes.length > attachmentMaxBytes) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " downloaded to " + bytes.length
                            + " bytes, above the " + attachmentMaxBytes + " byte inline cap")
                    .asRuntimeException();
        }
        return item.toBuilder().setContent(ByteString.copyFrom(bytes)).build();
    }

    private static <T extends Message> T requireValid(T message) {
        VALIDATOR.requireValid(message);
        return message;
    }

    private static void fail(StreamObserver<?> observer, Throwable t) {
        if (t instanceof io.grpc.StatusRuntimeException e) {
            observer.onError(e);
            return;
        }
        Status status;
        if (t instanceof GraphClient.GraphApiException e) {
            status = (switch (e.status()) {
                case 400 -> Status.INVALID_ARGUMENT;
                case 401, 403 -> Status.PERMISSION_DENIED;
                case 404 -> Status.NOT_FOUND;
                case 429 -> Status.RESOURCE_EXHAUSTED;
                default -> Status.INTERNAL;
            }).withDescription(e.getMessage());
        } else if (t instanceof InterruptedException e) {
            Thread.currentThread().interrupt();
            status = Status.CANCELLED.withDescription("interrupted: " + e.getMessage());
        } else if (t instanceof IOException e) {
            status = Status.UNAVAILABLE.withDescription("Graph unreachable: " + e.getMessage());
        } else {
            status = Status.INTERNAL.withDescription(String.valueOf(t.getMessage()));
        }
        observer.onError(status.asRuntimeException());
    }
}
