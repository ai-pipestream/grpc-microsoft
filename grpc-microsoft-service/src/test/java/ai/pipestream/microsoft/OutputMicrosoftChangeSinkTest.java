package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.output.MicrosoftConnectorOutputFormat;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import org.junit.jupiter.api.Test;

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
}
