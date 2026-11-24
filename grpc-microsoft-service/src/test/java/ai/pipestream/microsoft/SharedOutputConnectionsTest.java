package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.output.OutputFormat;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import com.google.protobuf.Message;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The process has one output sink, and every connection's Sync writes
 * through it. A full sweep of one connection says nothing about the others:
 * their records stay in the bundle. And two Syncs running at once each
 * complete with their own records, not with whatever the other had emitted
 * so far.
 */
class SharedOutputConnectionsTest {

    private final List<String> completes = new ArrayList<>();
    private final OutputMicrosoftChangeSink sink =
            new OutputMicrosoftChangeSink(nullStore(), List.of(recorder()), "out");

    @Test
    void aFullSweepOfOneConnectionKeepsTheRecordsOfAnother() {
        MicrosoftChangeSink x = sink.forConnection("x");
        MicrosoftChangeSink y = sink.forConnection("y");

        x.emit(file("x-drive", "x1"));
        x.completeRun("run-x", MicrosoftChangeSink.ALL_KINDS);
        y.emit(file("y-drive", "y1"));
        y.completeRun("run-y", MicrosoftChangeSink.ALL_KINDS);

        assertThat(completes.get(completes.size() - 1))
                .as("the bundle after connection y's full sweep")
                .isEqualTo("out=[x1, y1]");
    }

    @Test
    void twoSyncsAtOnceDoNotCompleteWithEachOthersRecords() {
        MicrosoftChangeSink x = sink.forConnection("x");
        MicrosoftChangeSink y = sink.forConnection("y");

        // Both passes are under way when x finishes.
        x.emit(file("x-drive", "x1"));
        y.emit(file("y-drive", "y1"));
        x.completeRun("run-x", MicrosoftChangeSink.ALL_KINDS);
        y.emit(file("y-drive", "y2"));
        y.completeRun("run-y", MicrosoftChangeSink.ALL_KINDS);

        assertThat(completes).as("x's bundle holds x's records, then y's pass adds all of its own")
                .containsExactly("out=[x1]", "out=[x1, y1, y2]");
    }

    @Test
    void anIncrementOfAConnectionIsWrittenUnderItsOwnName() {
        MicrosoftChangeSink x = sink.forConnection("x");
        MicrosoftChangeSink y = sink.forConnection("tenant/y");

        x.emit(file("x-drive", "x1"));
        x.completeRun("run-x", MicrosoftChangeSink.ALL_KINDS);
        y.emit(file("y-drive", "y1"));
        y.completeRun("run-y", java.util.Set.of());
        sink.emit(file("d", "plain"));
        sink.completeRun("run-plain", java.util.Set.of());

        assertThat(completes).containsExactly("out=[x1]",
                "out/increments/" + ai.pipestream.output.ObjectKeys.segment("tenant/y")
                        + "/run-y=[y1]",
                "out/increments/run-plain=[plain]");
    }

    @Test
    void theOkfBundleKeepsAnotherConnectionsConcepts(@org.junit.jupiter.api.io.TempDir
            java.nio.file.Path dir) {
        OkfMicrosoftChangeSink okf = new OkfMicrosoftChangeSink(new ai.pipestream.okf.OkfOutput(
                dir.resolve("okf"), dir.resolve("okf.zip"), dir.resolve("okf.warc.gz")), null);
        MicrosoftChangeSink x = okf.forConnection("x");
        MicrosoftChangeSink y = okf.forConnection("y");

        x.emit(named("x-drive", "x1"));
        y.emit(named("y-drive", "y1"));
        x.completeRun("run-x", MicrosoftChangeSink.ALL_KINDS);
        assertThat(dir.resolve("okf/items/x-drive/x1.md")).exists();
        assertThat(dir.resolve("okf/items/y-drive/y1.md"))
                .as("y's concept before y's own pass completed").doesNotExist();
        y.completeRun("run-y", MicrosoftChangeSink.ALL_KINDS);

        assertThat(dir.resolve("okf/items/x-drive/x1.md"))
                .as("x's concept after y's full sweep").exists();
        assertThat(dir.resolve("okf/items/y-drive/y1.md")).exists();
    }

    private static MicrosoftChange named(String driveId, String id) {
        MicrosoftChange change = file(driveId, id);
        return change.toBuilder().setEntity(change.getEntity().toBuilder()
                .setIngestedAt(com.google.protobuf.Timestamp.newBuilder().setSeconds(1))
                .setDriveItem(change.getEntity().getDriveItem().toBuilder().setName(id + ".txt")
                        .setWebUrl("https://contoso.sharepoint.com/" + id + ".txt"))).build();
    }

    private static MicrosoftChange file(String driveId, String id) {
        return MicrosoftChange.newBuilder().setChangeId(id)
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder().setEntityId(id)
                        .setDriveItem(DriveItem.newBuilder().setId(id).setDriveId(driveId)))
                .build();
    }

    private OutputFormat recorder() {
        return new OutputFormat() {
            @Override
            public String id() {
                return "recorder";
            }

            @Override
            public boolean supports(Message record) {
                return false;
            }

            @Override
            public void complete(OutputStore store, List<Message> records, String prefix) {
                completes.add(prefix + "=" + records.stream()
                        .map(r -> ((MicrosoftChange) r).getEntity().getEntityId()).sorted().toList());
            }
        };
    }

    private static OutputStore nullStore() {
        return new OutputStore() {
            @Override
            public String id() {
                return "null";
            }

            @Override
            public boolean available(Map<String, String> env) {
                return true;
            }

            @Override
            public void open(Map<String, String> env) {
            }

            @Override
            public void put(OutputObject object) {
            }
        };
    }
}
