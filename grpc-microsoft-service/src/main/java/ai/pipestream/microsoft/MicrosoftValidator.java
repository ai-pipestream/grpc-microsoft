package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.Site;
import com.google.protobuf.Message;
import com.google.protobuf.Timestamp;
import io.grpc.Status;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * Programmatic identity and numeric-floor rules for the Microsoft Graph
 * domain model. Applied before an entity leaves the mapper, crawler, or
 * gRPC facade.
 */
public final class MicrosoftValidator {

    private static final MicrosoftValidator INSTANCE = new MicrosoftValidator();
    private static final Pattern EMAIL = Pattern.compile("^[^ @]+@[^ @]+[.][^ @]+$");

    private MicrosoftValidator() {
    }

    /**
     * Returns the shared validator.
     *
     * @return the singleton instance
     */
    public static MicrosoftValidator create() {
        return INSTANCE;
    }

    /**
     * One rule failure: proto field path, rule id, and human message.
     *
     * @param path the proto field path, empty for the message itself
     * @param ruleId a stable rule identifier
     * @param message a human-readable failure
     */
    public record Violation(String path, String ruleId, String message) {
    }

    /**
     * The collected violations for one message (empty = valid).
     *
     * @param violations the failures; copied
     */
    public record ValidationResult(List<Violation> violations) {
        /** Validates and normalizes fields. */
        public ValidationResult {
            violations = List.copyOf(violations);
        }

        /**
         * Whether the message passed every rule.
         *
         * @return {@code true} when {@code violations} is empty
         */
        public boolean isValid() {
            return violations.isEmpty();
        }

        /** Throws {@code FAILED_PRECONDITION} when this result is not valid. */
        public void throwIfInvalid() {
            if (violations.isEmpty()) {
                return;
            }
            StringBuilder sb = new StringBuilder("microsoft validation failed:");
            for (Violation violation : violations) {
                sb.append(" [").append(violation.ruleId()).append(" at ")
                        .append(violation.path().isEmpty() ? "<message>" : violation.path())
                        .append(": ").append(violation.message()).append(']');
            }
            throw Status.FAILED_PRECONDITION.withDescription(sb.toString()).asRuntimeException();
        }
    }

    /**
     * Checks {@code message} against the domain rules.
     *
     * @param message a Microsoft Graph domain proto
     * @return the collected violations (empty when valid)
     */
    public ValidationResult validate(Message message) {
        List<Violation> violations = new ArrayList<>();
        visit("", message, violations);
        return new ValidationResult(violations);
    }

    /**
     * Validates {@code message} and throws if any rule fails.
     *
     * @param message a Microsoft Graph domain proto
     */
    public void requireValid(Message message) {
        validate(message).throwIfInvalid();
    }

    private static void visit(String prefix, Message message, List<Violation> out) {
        switch (message) {
            case Site site -> required(prefix, "id", site.getId(), out);
            case Drive drive -> required(prefix, "id", drive.getId(), out);
            case DriveItem item -> visitDriveItem(prefix, item, out);
            case GraphUser user -> visitUser(prefix, user, out);
            case MicrosoftEntity entity -> visitEntity(prefix, entity, out);
            case MicrosoftSnapshot snapshot -> required(prefix, "snapshot_id", snapshot.getSnapshotId(), out);
            case MicrosoftChange change -> visitChange(prefix, change, out);
            default -> {
            }
        }
    }

    private static void visitDriveItem(String prefix, DriveItem item, List<Violation> out) {
        required(prefix, "id", item.getId(), out);
        if (item.getSize() < 0) {
            out.add(new Violation(child(prefix, "size"), "int64.gte", "size must be >= 0"));
        }
    }

    private static void visitUser(String prefix, GraphUser user, List<Violation> out) {
        required(prefix, "id", user.getId(), out);
        if (!user.getMail().isEmpty() && !EMAIL.matcher(user.getMail()).matches()) {
            out.add(new Violation(child(prefix, "mail"), "user.mail_format",
                    "mail must be empty or a valid address"));
        }
    }

    private static void visitEntity(String prefix, MicrosoftEntity entity, List<Violation> out) {
        required(prefix, "entity_id", entity.getEntityId(), out);
        requiredTimestamp(prefix, "ingested_at", entity.hasIngestedAt(), entity.getIngestedAt(), out);
        if (entity.getEntityCase() != MicrosoftEntity.EntityCase.ENTITY_NOT_SET) {
            Message nested = switch (entity.getEntityCase()) {
                case SITE -> entity.getSite();
                case DRIVE -> entity.getDrive();
                case DRIVE_ITEM -> entity.getDriveItem();
                case USER -> entity.getUser();
                case ENTITY_NOT_SET -> entity;
            };
            visit(child(prefix, entity.getEntityCase().name().toLowerCase()), nested, out);
        }
    }

    private static void visitChange(String prefix, MicrosoftChange change, List<Violation> out) {
        required(prefix, "change_id", change.getChangeId(), out);
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_UPSERT && !change.hasEntity()) {
            out.add(new Violation(prefix, "change.upsert_has_entity",
                    "an UPSERT change must carry the entity it upserts"));
        }
        if (change.hasEntity()) {
            visitEntity(child(prefix, "entity"), change.getEntity(), out);
        }
    }

    private static void required(String prefix, String field, String value, List<Violation> out) {
        if (value == null || value.isEmpty()) {
            out.add(new Violation(child(prefix, field), "required", field + " is required"));
        }
    }

    private static void requiredTimestamp(String prefix, String field, boolean present, Timestamp value,
            List<Violation> out) {
        if (!present || (value.getSeconds() == 0 && value.getNanos() == 0)) {
            out.add(new Violation(child(prefix, field), "required", field + " is required"));
        }
    }

    private static String child(String prefix, String field) {
        return prefix == null || prefix.isEmpty() ? field : prefix + "." + field;
    }
}
