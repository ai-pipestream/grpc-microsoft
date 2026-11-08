package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.output.OutputFormats;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import com.google.protobuf.StringValue;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftConnectorOutputFormatTest {

    @Test
    void writesConnectorProtobufForMicrosoftChangesOnly() throws Exception {
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
        MicrosoftChange file = MicrosoftChange.newBuilder()
                .setChangeId("m")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setDriveItem(DriveItem.newBuilder().setId("file-1").setDriveId("d")))
                .build();
        MicrosoftConnectorOutputFormat connector = new MicrosoftConnectorOutputFormat();
        assertThat(connector.supports(StringValue.of("other"))).isFalse();
        assertThat(connector.supports(file)).isTrue();
        connector.emit(store, file, "drives/d/items/file-1");
        assertThat(objects).extracting(OutputObject::key)
                .containsExactly("drives/d/items/file-1.connector.pb");
        assertThat(objects.get(0).content()).isEqualTo(file.toByteArray());
    }

    @Test
    void registersWithServiceLoader() {
        assertThat(OutputFormats.load().has(MicrosoftConnectorOutputFormat.ID)).isTrue();
    }
}
