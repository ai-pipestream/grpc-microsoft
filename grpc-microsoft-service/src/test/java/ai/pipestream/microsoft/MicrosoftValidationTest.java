package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import com.google.protobuf.Message;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftValidationTest {

    private static final MicrosoftValidator VALIDATOR = MicrosoftValidator.create();

    private static void assertViolation(Message message, String path, String ruleId) {
        assertThat(VALIDATOR.validate(message).violations())
                .as("expected %s at %s", ruleId, path)
                .anyMatch(v -> v.path().equals(path) && v.ruleId().equals(ruleId));
    }

    @Test
    void identityAndNumericFloors() {
        assertViolation(DriveItem.getDefaultInstance(), "id", "required");
        assertViolation(DriveItem.newBuilder().setId("i").setSize(-1).build(), "size", "int64.gte");
        assertThat(VALIDATOR.validate(DriveItem.newBuilder().setId("i").build()).violations())
                .isEmpty();
    }

    @Test
    void userMailWhenPresent() {
        GraphUser valid = GraphUser.newBuilder().setId("u1").build();
        assertThat(VALIDATOR.validate(valid).violations()).isEmpty();
        assertThat(VALIDATOR.validate(valid.toBuilder().setMail("bot@contoso.com").build())
                .violations()).isEmpty();
        assertViolation(valid.toBuilder().setMail("not-an-email").build(), "mail",
                "user.mail_format");
    }

    @Test
    void upsertMustCarryEntity() {
        MicrosoftChange delete = MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .build();
        assertThat(VALIDATOR.validate(delete).violations()).isEmpty();
        assertViolation(delete.toBuilder()
                        .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                        .build(),
                "", "change.upsert_has_entity");
        assertThat(VALIDATOR.validate(MicrosoftChange.newBuilder()
                .setChangeId("c2")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1)))
                .build()).violations()).isEmpty();
    }

    @Test
    void snapshotNeedsId() {
        assertViolation(MicrosoftSnapshot.getDefaultInstance(), "snapshot_id", "required");
    }
}
