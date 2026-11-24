package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.DegradedProjection;
import ai.pipestream.microsoft.v1.DegradedProjections;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemInclude;
import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.ManagedChannel;
import io.grpc.Metadata;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.MetadataUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.catchThrowableOfType;

/**
 * A read asks Graph for an explicit property list, and may be answered from
 * Graph's default one when Graph rejects the list. That is only right when
 * the query options are what Graph rejected, and it is never silent: the
 * fields only the list asks for ({@code sharepointIds}, {@code publication},
 * the download URL) are then missing from every entity, so the caller is
 * told which route was answered without them.
 */
class ProjectionFallbackTest {

    /** The trailer a call ends with when a read was answered from Graph's default properties. */
    private static final Metadata.Key<byte[]> DEGRADED = Metadata.Key.of(
            "ai.pipestream.microsoft.v1.DegradedProjections-bin", Metadata.BINARY_BYTE_MARSHALLER);

    private static final String SELECT_REJECTED = "Parsing OData Select and Expand failed: "
            + "Could not find a property named 'sharepointIds' on type 'microsoft.graph.driveItem'.";

    private static final String SELECTED = "/drives/d1/items/x1?$select=" + GraphPaths.DRIVE_ITEM_SELECT;

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private final AtomicReference<Metadata> trailers = new AtomicReference<>();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token",
                AuthMode.AUTH_MODE_UNSPECIFIED));
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, files, 0, null, null, 4))
                .build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel).withInterceptors(
                MetadataUtils.newCaptureMetadataInterceptor(new AtomicReference<>(), trailers));
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    @Test
    void aBadRequestThatIsNotAboutTheQueryIsNotAnsweredWithoutIt() {
        fake.stub(SELECTED, error(400, "invalidRequest", "The provided item id is not valid"));
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));

        StatusRuntimeException failure = catchThrowableOfType(StatusRuntimeException.class,
                () -> stub.getItem(request()));

        assertThat(failure).as("GetItem whose request Graph rejected for another reason").isNotNull();
        assertThat(failure.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(failure.getStatus().getDescription()).contains("The provided item id is not valid");
        assertThat(fake.requestsTo("/drives/d1/items/x1")).extracting(r -> r.query())
                .as("the request is not sent again with the properties dropped")
                .containsExactly("$select=" + GraphPaths.DRIVE_ITEM_SELECT);
    }

    @Test
    void aReadAnsweredFromGraphDefaultPropertiesSaysSoToTheCaller() {
        fake.stub(SELECTED, error(400, "BadRequest", SELECT_REJECTED));
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));

        DriveItem item = stub.getItem(request()).getItem();

        assertThat(item.getName()).isEqualTo("a.txt");
        assertThat(trailers.get().get(DEGRADED))
                .as("the trailer naming the read that was answered without its properties")
                .isNotNull();
    }

    @Test
    void aReadThatGraphAnswersInFullCarriesNoSuchTrailer() {
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));

        stub.getItem(request());

        assertThat(trailers.get().get(DEGRADED)).isNull();
    }

    @Test
    void theTrailerNamesTheRouteTheQueryAndWhatGraphSaid() {
        fake.stub(SELECTED, error(400, "BadRequest", SELECT_REJECTED));
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));

        stub.getItem(request());

        DegradedProjections degraded = trailers.get().get(ProjectionReport.TRAILER);
        assertThat(degraded.getOmitted()).isZero();
        assertThat(degraded.getProjectionsList()).singleElement().satisfies(read -> {
            assertThat(read.getRoute()).isEqualTo("/drives/d1/items/x1");
            assertThat(read.getRejectedQuery()).startsWith("$select=id,name,")
                    .hasSizeLessThanOrEqualTo(ProjectionReport.MAX_TEXT);
            assertThat(read.getUsedQuery()).isEmpty();
            assertThat(read.getGraphCode()).isEqualTo("BadRequest");
            assertThat(read.getGraphMessage()).contains("sharepointIds");
            assertThat(read.getReads()).isEqualTo(1);
        });
    }

    @Test
    void aFailedCallStillSaysWhichReadsWereDegraded() {
        fake.stub(SELECTED, error(400, "BadRequest", SELECT_REJECTED));
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));
        fake.stub("/drives/d1/items/x1/children?$select=" + GraphPaths.DRIVE_ITEM_SELECT,
                error(400, "BadRequest", SELECT_REJECTED));
        fake.stub("/drives/d1/items/x1/children", error(403, "accessDenied", "no"));

        StatusRuntimeException failure = catchThrowableOfType(StatusRuntimeException.class,
                () -> stub.listChildren(ListChildrenRequest.newBuilder().setDriveId("d1")
                        .setFolderId("x1").build()).forEachRemaining(r -> { }));

        assertThat(failure).isNotNull();
        assertThat(failure.getStatus().getCode()).isEqualTo(Status.Code.PERMISSION_DENIED);
        // The children read never got an answer, so nothing was answered degraded.
        assertThat(trailers.get().get(ProjectionReport.TRAILER)).isNull();
    }

    @Test
    void aSyncCountsEveryReadOfAKindOnceInTheTrailer() {
        for (String drive : List.of("d1", "d2")) {
            fake.stub("/drives/" + drive + "?$select=" + GraphPaths.DRIVE_SELECT,
                    error(400, "invalidRequest", "Invalid $select: the system facet is not supported"));
            fake.stub("/drives/" + drive, MicrosoftFixtures.driveJson(drive, "Docs " + drive));
            fake.stub("/drives/" + drive + "/root/delta?$select=" + GraphPaths.driveItemSelect(false),
                    error(400, "BadRequest", SELECT_REJECTED));
            fake.stub("/drives/" + drive + "/root/delta", MicrosoftFixtures.deltaJson(
                    fake.baseUrl() + "/drives/" + drive + "/root/delta?token=t1",
                    MicrosoftFixtures.fileJson("f-" + drive, "a.txt", drive)));
        }

        List<SyncResponse> responses = new ArrayList<>();
        stub.sync(SyncRequest.newBuilder().addDriveIds("d1").addDriveIds("d2").build())
                .forEachRemaining(responses::add);

        assertThat(responses.get(responses.size() - 1).getEventCase())
                .isEqualTo(SyncResponse.EventCase.RESUME_CURSOR);
        DegradedProjections degraded = trailers.get().get(ProjectionReport.TRAILER);
        assertThat(degraded.getProjectionsList()).extracting(DegradedProjection::getRoute,
                DegradedProjection::getReads).containsExactly(
                org.assertj.core.groups.Tuple.tuple("/drives/d1", 2L),
                org.assertj.core.groups.Tuple.tuple("/drives/d1/root/delta", 2L));
    }

    @Test
    void onlyAFewDistinctReadsAreListedAndTheRestAreCounted() {
        ProjectionReport report = new ProjectionReport();
        GraphClient.GraphApiException rejection = new GraphClient.GraphApiException(400,
                "BadRequest", SELECT_REJECTED, "x".repeat(5000));

        io.grpc.Context.current().withValue(ProjectionReport.CURRENT, report).run(() -> {
            for (int i = 0; i < ProjectionReport.MAX_LISTED + 3; i++) {
                ProjectionReport.degraded("/route/" + i + "?$select=" + "p".repeat(1000) + i,
                        "/route/" + i, rejection);
            }
        });

        DegradedProjections degraded = report.build();
        assertThat(degraded.getProjectionsCount()).isEqualTo(ProjectionReport.MAX_LISTED);
        assertThat(degraded.getOmitted()).isEqualTo(3);
        // Trailers share the client's 8 KiB metadata limit.
        assertThat(degraded.getSerializedSize()).isLessThan(4096);
    }

    @Test
    void onlyARejectionOfTheQueryOptionsIsOne() {
        assertThat(rejection(400, "BadRequest", SELECT_REJECTED).rejectsQueryOptions()).isTrue();
        assertThat(rejection(400, "invalidRequest", "Invalid $expand").rejectsQueryOptions()).isTrue();
        assertThat(rejection(400, "BadRequest", "Query option 'Select' is not allowed.")
                .rejectsQueryOptions()).isTrue();
        assertThat(rejection(400, "invalidRequest", "The provided item id is not valid")
                .rejectsQueryOptions()).isFalse();
        assertThat(rejection(400, "BadRequest", "BadRequest").rejectsQueryOptions()).isFalse();
        assertThat(rejection(400, "invalidRequest", "The selected range is expanded")
                .rejectsQueryOptions()).isFalse();
        assertThat(rejection(403, "accessDenied", "$select is not allowed for this caller")
                .rejectsQueryOptions()).isFalse();
    }

    @Test
    void theRouteAndTheQueryOfACandidateAreToldApart() {
        assertThat(ProjectionReport.route("/drives/d1/items/x1?$select=id")).isEqualTo("/drives/d1/items/x1");
        assertThat(ProjectionReport.query("/drives/d1/items/x1?$select=id")).isEqualTo("$select=id");
        assertThat(ProjectionReport.route("https://graph.example/v1.0/me/drives?$top=1"))
                .isEqualTo("/v1.0/me/drives");
        assertThat(ProjectionReport.query("/me/drives")).isEmpty();
    }

    private static GraphClient.GraphApiException rejection(int status, String code, String message) {
        return new GraphClient.GraphApiException(status, code, message,
                "{\"error\":{\"code\":\"" + code + "\",\"message\":\"" + message + "\"}}");
    }

    private static GetItemRequest request() {
        return GetItemRequest.newBuilder().setDriveId("d1").setItemId("x1")
                .setInclude(DriveItemInclude.getDefaultInstance()).build();
    }

    private static FakeGraphServer.Stub error(int status, String code, String message) {
        return new FakeGraphServer.Stub(status, ("{\"error\":{\"code\":\"" + code
                + "\",\"message\":\"" + message + "\"}}").getBytes(StandardCharsets.UTF_8),
                Map.of("content-type", "application/json"));
    }
}
