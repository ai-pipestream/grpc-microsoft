package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.FieldWriteState;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import ai.pipestream.microsoft.v1.UploadProgress;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.ByteString;
import io.grpc.Status;
import io.grpc.stub.StreamObserver;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Base64;
import java.util.HexFormat;
import java.util.Map;

/** Receives the frames of one {@code UploadItem} call and stores the file through Graph. */
final class UploadHandler implements StreamObserver<UploadItemRequest> {

    private static final byte[] PROCESS_KEY = new byte[32];

    static {
        new SecureRandom().nextBytes(PROCESS_KEY);
    }

    private final MicrosoftSessions sessions;
    private final MicrosoftMapper mapper;
    private final WritePlans plans;
    private final StreamObserver<UploadItemResponse> responses;

    private MicrosoftSessions.Session session;
    private UploadItemHeader header;
    private String itemRoute;
    private MessageDigest digest;
    private long received;
    private boolean resumed;
    private boolean closed;

    /** Content of a small upload of known size. */
    private byte[] small;
    /** The upload session of a large upload of known size. */
    private SessionUpload upload;
    /** Where content of unknown size is kept until the end frame. */
    private Path spool;
    private OutputStream spoolOut;

    UploadHandler(MicrosoftSessions sessions, MicrosoftMapper mapper, WritePlans plans,
            StreamObserver<UploadItemResponse> responses) {
        this.sessions = sessions;
        this.mapper = mapper;
        this.plans = plans;
        this.responses = responses;
    }

    @Override
    public void onNext(UploadItemRequest frame) {
        if (closed) {
            return;
        }
        try {
            switch (frame.getFrameCase()) {
                case HEADER -> start(frame.getHeader());
                case CHUNK -> accept(frame.getChunk());
                case END -> finish(frame.getEnd());
                case FRAME_NOT_SET -> throw new IllegalArgumentException("empty upload frame");
            }
        } catch (Throwable t) {
            fail(t);
        }
    }

    @Override
    public void onError(Throwable t) {
        closed = true;
        discardSpool();
    }

    @Override
    public void onCompleted() {
        if (!closed) {
            fail(new IllegalArgumentException("the upload ended without an end frame"));
        }
    }

    private void fail(Throwable t) {
        closed = true;
        discardSpool();
        responses.onError(GraphErrors.toException(t));
    }

    private void start(UploadItemHeader first) throws Exception {
        if (header != null) {
            throw new IllegalArgumentException("the upload has two headers");
        }
        if (first.getResumeOffset() < 0 || first.getFragmentBytes() < 0
                || (first.hasSize() && first.getSize() < 0)) {
            throw new IllegalArgumentException(
                    "size, resume_offset and fragment_bytes must not be negative");
        }
        session = sessions.resolve(first.getConnectionId());
        itemRoute = GraphPaths.childByName(first.getDriveId(), first.getParentId(),
                first.getParentPath(), first.getName());
        WritePlans.conflict(first.getConflictBehavior());
        header = first;
        digest = MessageDigest.getInstance("SHA-256");
        if (!first.getResumeToken().isEmpty()) {
            resume();
        } else if (!first.hasSize()) {
            spool = Files.createTempFile("upload-", ".part");
            spoolOut = Files.newOutputStream(spool);
        } else if (first.getSize() <= GraphFiles.SIMPLE_UPLOAD_LIMIT) {
            small = new byte[(int) first.getSize()];
        } else {
            upload = open(first.getSize());
        }
    }

    private void resume() throws Exception {
        if (!header.hasSize()) {
            throw new IllegalArgumentException("a resumed upload needs size");
        }
        String uploadUrl = unwrap(header.getResumeToken());
        long expected = session.files().uploadSessionOffset(uploadUrl);
        if (expected >= header.getSize()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("the upload session has no bytes left to receive")
                    .asRuntimeException();
        }
        if (expected != header.getResumeOffset()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("the upload session expects offset " + expected)
                    .asRuntimeException();
        }
        resumed = true;
        received = expected;
        upload = new SessionUpload(uploadUrl, header.getSize(), expected);
        progress(expected, uploadUrl);
    }

    private SessionUpload open(long total) throws Exception {
        String uploadUrl = session.files().createUploadSession(itemRoute, header.getName(),
                WritePlans.conflict(header.getConflictBehavior()), header.getExpectedEtag());
        progress(0, uploadUrl);
        return new SessionUpload(uploadUrl, total, 0);
    }

    private void accept(ByteString chunk) throws Exception {
        requireHeader();
        if (header.hasSize() && received + chunk.size() > header.getSize()) {
            throw new IllegalArgumentException("the content is longer than the declared size of "
                    + header.getSize() + " bytes");
        }
        for (java.nio.ByteBuffer part : chunk.asReadOnlyByteBufferList()) {
            digest.update(part.duplicate());
        }
        if (small != null) {
            chunk.copyTo(small, (int) received);
        } else if (upload != null) {
            for (ByteString part = chunk; !part.isEmpty();) {
                part = part.substring(upload.write(part));
            }
        } else {
            chunk.writeTo(spoolOut);
        }
        received += chunk.size();
    }

    private void finish(UploadItemEnd end) throws Exception {
        requireHeader();
        if (header.hasSize() && received != header.getSize()) {
            throw new IllegalArgumentException("the content is " + received
                    + " bytes, not the declared " + header.getSize());
        }
        String expected = end.getSha256().trim();
        if (!expected.isEmpty()) {
            if (resumed) {
                throw new IllegalArgumentException("sha256 cannot be verified on a resumed upload");
            }
            String actual = HexFormat.of().formatHex(digest.digest());
            if (!actual.equalsIgnoreCase(expected)) {
                if (upload != null) {
                    session.files().cancelUploadSession(upload.uploadUrl);
                }
                throw Status.DATA_LOSS
                        .withDescription("sha256 of the received content is " + actual)
                        .asRuntimeException();
            }
        }
        JsonNode stored;
        if (small != null) {
            stored = putSmall(small, small.length);
        } else if (upload != null) {
            stored = upload.finish();
        } else {
            stored = sendSpool();
        }
        DriveItem item = mapper.toDriveItem(stored, header.getDriveId());
        MicrosoftValidator.create().requireValid(item);
        if (header.getFieldsCount() > 0) {
            item = setFields(item);
        }
        closed = true;
        responses.onNext(UploadItemResponse.newBuilder().setItem(item).build());
        responses.onCompleted();
    }

    private JsonNode putSmall(byte[] content, int length) throws Exception {
        return session.files().putContent(itemRoute, content, length, header.getMediaType(),
                WritePlans.conflict(header.getConflictBehavior()), header.getExpectedEtag());
    }

    private JsonNode sendSpool() throws Exception {
        spoolOut.close();
        try (InputStream in = Files.newInputStream(spool)) {
            if (received <= GraphFiles.SIMPLE_UPLOAD_LIMIT) {
                return putSmall(in.readAllBytes(), (int) received);
            }
            upload = open(received);
            byte[] buffer = new byte[64 * 1024];
            for (int read = in.read(buffer); read >= 0; read = in.read(buffer)) {
                for (int offset = 0; offset < read;) {
                    offset += upload.write(ByteString.copyFrom(buffer, offset, read - offset));
                }
            }
            return upload.finish();
        } finally {
            discardSpool();
        }
    }

    private DriveItem setFields(DriveItem item) throws Exception {
        ListFieldWriter.Patch patch = ListFieldWriter.translate(header.getFieldsList(),
                plans.driveLookups(session, header.getDriveId()),
                FieldWriteState.FIELD_WRITE_STATE_WRITTEN);
        String route = GraphPaths.item(header.getDriveId(), item.getId());
        JsonNode fields = session.graph().execute(new GraphClient.Call("PATCH",
                route + "/listItem/fields", Map.of(), patch.fields())).body();
        DriveItem current = mapper.toDriveItem(session.graph().get(route, Map.of()),
                header.getDriveId());
        MicrosoftValidator.create().requireValid(current);
        return mapper.withListColumns(current, fields);
    }

    private void requireHeader() {
        if (header == null) {
            throw new IllegalArgumentException("the first upload frame must be the header");
        }
    }

    private void progress(long committed, String uploadUrl) {
        responses.onNext(UploadItemResponse.newBuilder()
                .setProgress(UploadProgress.newBuilder().setCommittedBytes(committed)
                        .setResumeToken(wrap(uploadUrl)))
                .build());
    }

    private void discardSpool() {
        try {
            if (spoolOut != null) {
                spoolOut.close();
            }
            if (spool != null) {
                Files.deleteIfExists(spool);
            }
        } catch (IOException ignored) {
            // the temporary file is left for the operating system to remove
        }
        spool = null;
        spoolOut = null;
    }

    /** The request size for upload sessions: a positive multiple of 320 KiB up to 60 MiB. */
    static int fragmentBytes(int requested) {
        int wanted = requested == 0 ? GraphFiles.DEFAULT_FRAGMENT_BYTES
                : Math.min(requested, GraphFiles.MAX_FRAGMENT_BYTES);
        int units = Math.max(1, wanted / GraphFiles.FRAGMENT_UNIT_BYTES);
        return units * GraphFiles.FRAGMENT_UNIT_BYTES;
    }

    private String wrap(String uploadUrl) {
        Base64.Encoder encoder = Base64.getUrlEncoder().withoutPadding();
        return encoder.encodeToString(uploadUrl.getBytes(StandardCharsets.UTF_8)) + "."
                + encoder.encodeToString(sign(uploadUrl));
    }

    private String unwrap(String token) {
        int dot = token.indexOf('.');
        try {
            if (dot > 0) {
                String uploadUrl = new String(Base64.getUrlDecoder().decode(token.substring(0, dot)),
                        StandardCharsets.UTF_8);
                byte[] signature = Base64.getUrlDecoder().decode(token.substring(dot + 1));
                if (MessageDigest.isEqual(signature, sign(uploadUrl))) {
                    return uploadUrl;
                }
            }
        } catch (IllegalArgumentException e) {
            // not base64: rejected below
        }
        throw new IllegalArgumentException("resume_token was not issued for this connection");
    }

    private byte[] sign(String uploadUrl) {
        String secret = session.config().clientSecret();
        byte[] key = secret == null || secret.isEmpty() ? PROCESS_KEY
                : secret.getBytes(StandardCharsets.UTF_8);
        try {
            Mac mac = Mac.getInstance("HmacSHA256");
            mac.init(new SecretKeySpec(key, "HmacSHA256"));
            return mac.doFinal(uploadUrl.getBytes(StandardCharsets.UTF_8));
        } catch (GeneralSecurityException e) {
            throw new IllegalStateException(e);
        }
    }

    /** One upload session: buffers content and sends it a fragment at a time. */
    private final class SessionUpload {

        private final String uploadUrl;
        private final long total;
        private final byte[] buffer;
        private long committed;
        private int filled;

        SessionUpload(String uploadUrl, long total, long committed) {
            this.uploadUrl = uploadUrl;
            this.total = total;
            this.committed = committed;
            this.buffer = new byte[(int) Math.min(fragmentBytes(header.getFragmentBytes()),
                    Math.max(1L, total - committed))];
        }

        /** Takes bytes from the front of {@code data}; returns how many were taken. */
        int write(ByteString data) throws Exception {
            if (filled == buffer.length) {
                long sent = committed + filled;
                committed = GraphFiles.nextExpectedOffset(send(), sent);
                if (committed != sent) {
                    throw new IOException("the upload session expects offset " + committed
                            + " after " + sent + " bytes were sent");
                }
                filled = 0;
                progress(committed, uploadUrl);
            }
            int taken = Math.min(data.size(), buffer.length - filled);
            data.substring(0, taken).copyTo(buffer, filled);
            filled += taken;
            return taken;
        }

        /** Sends what is left and returns the stored driveItem. */
        JsonNode finish() throws Exception {
            JsonNode stored = send();
            progress(total, uploadUrl);
            return stored;
        }

        private JsonNode send() throws Exception {
            return session.files().putFragment(uploadUrl, buffer, filled, committed, total);
        }
    }
}
