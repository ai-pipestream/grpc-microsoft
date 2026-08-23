package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.okf.OkfOutput;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class OkfMicrosoftChangeSinkTest {

    private FakeGraphServer fake;

    @AfterEach
    void stop() {
        if (fake != null) {
            fake.close();
        }
    }

    @Test
    void completeRunWritesLocalArtifacts(@TempDir Path dir) {
        OkfMicrosoftChangeSink sink = new OkfMicrosoftChangeSink(new OkfOutput(
                dir.resolve("okf"), dir.resolve("okf.zip"), dir.resolve("okf.warc.gz")),
                null);
        sink.emit(MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                        .setDriveItem(DriveItem.newBuilder()
                                .setId("file-1")
                                .setName("notes.txt")
                                .setDriveId("drive-1")
                                .setWebUrl("https://contoso.sharepoint.com/notes.txt")))
                .build());
        sink.completeRun("run-1");
        assertThat(dir.resolve("okf/items/drive-1/file-1.md")).exists();
        assertThat(dir.resolve("okf.zip")).exists();
        assertThat(dir.resolve("okf.warc.gz")).exists();
    }

    @Test
    void uploadsTreeToSharePoint(@TempDir Path dir) throws Exception {
        fake = FakeGraphServer.start();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        SharePointOkfPublisher publisher = new SharePointOkfPublisher(files, "drive-1",
                "/Knowledge");
        OkfMicrosoftChangeSink sink = new OkfMicrosoftChangeSink(new OkfOutput(
                dir.resolve("okf"), dir.resolve("okf.zip"), dir.resolve("okf.warc.gz")),
                publisher);
        sink.emit(MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                        .setDriveItem(DriveItem.newBuilder()
                                .setId("file-1")
                                .setName("notes.txt")
                                .setDriveId("drive-1")
                                .setWebUrl("https://contoso.sharepoint.com/notes.txt")))
                .build());
        sink.completeRun("run-1");
        assertThat(fake.requests()).anyMatch(r -> "PUT".equals(r.method())
                && r.path().contains("/Knowledge/"));
        assertThat(Files.size(dir.resolve("okf.zip"))).isPositive();
    }
}
