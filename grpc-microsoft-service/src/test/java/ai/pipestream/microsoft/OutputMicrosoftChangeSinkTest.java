package ai.pipestream.microsoft;

import ai.pipestream.microsoft.output.MicrosoftConnectorOutputFormat;
import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.output.OutputFormat;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import org.junit.jupiter.api.Test;

import com.google.protobuf.Message;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class OutputMicrosoftChangeSinkTest {

    @Test
    void writesConnectorProtobufUnderDriveHierarchy() {
        List<OutputObject> objects = new ArrayList<>();
        OutputStore store = new OutputStore() {
            @Override
            public String id() {
                return "memory";
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
                objects.add(object);
            }
        };
        OutputMicrosoftChangeSink sink = new OutputMicrosoftChangeSink(store,
                List.of(new MicrosoftConnectorOutputFormat()), "");
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setDriveItem(DriveItem.newBuilder()
                                .setId("file-1")
                                .setDriveId("drive-1")
                                .setName("notes.txt")))
                .build();
        sink.emit(change);
        assertThat(objects).extracting(OutputObject::key)
                .containsExactly("drives/drive-1/items/file-1.connector.pb");
        assertThat(objects.get(0).content()).isEqualTo(change.toByteArray());
        sink.close();
    }

    @Test
    void incrementalRunsMergeIntoTheMirrorOrWriteAsideWithoutABaseline() {
        List<String> completes = new ArrayList<>();
        OutputFormat recorder = new OutputFormat() {
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
                        .map(r -> ((MicrosoftChange) r).getEntity().getEntityId()
                                + (((MicrosoftChange) r).getOperation()
                                        == ChangeOperation.CHANGE_OPERATION_DELETE ? "-" : ""))
                        .toList());
            }
        };
        OutputMicrosoftChangeSink sink = new OutputMicrosoftChangeSink(nullStore(), List.of(recorder), "out");

        sink.emit(file("a", ChangeOperation.CHANGE_OPERATION_UPSERT));
        sink.completeRun("inc0", false);
        sink.emit(file("a", ChangeOperation.CHANGE_OPERATION_UPSERT));
        sink.emit(file("b", ChangeOperation.CHANGE_OPERATION_UPSERT));
        sink.completeRun("full1", true);
        sink.emit(file("b", ChangeOperation.CHANGE_OPERATION_DELETE));
        sink.emit(file("c", ChangeOperation.CHANGE_OPERATION_UPSERT));
        sink.completeRun("inc2", false);

        assertThat(completes).containsExactly(
                "out/increments/inc0=[a]",
                "out=[a, b]",
                "out=[a, b-, c]");
    }

    private static MicrosoftChange file(String id, ChangeOperation op) {
        return MicrosoftChange.newBuilder().setChangeId(id).setOperation(op)
                .setEntity(MicrosoftEntity.newBuilder().setEntityId(id)
                        .setDriveItem(DriveItem.newBuilder().setId(id).setDriveId("d")))
                .build();
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
