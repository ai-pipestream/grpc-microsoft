package ai.pipestream.microsoft.connector;

/**
 * Stable identifier GCA stores on the connection. Changing this after a
 * connection exists fails that connection; do not reuse the sample UUID
 * from Microsoft's template.
 */
public final class ConnectorId {

    /**
     * Primary identifier for this connector. Same value belongs in the
     * GCA manifest when the connector is on-boarded.
     */
    public static final String VALUE = "43760992-c66e-46bd-b937-b63e021aa63b";

    private ConnectorId() {
    }
}
