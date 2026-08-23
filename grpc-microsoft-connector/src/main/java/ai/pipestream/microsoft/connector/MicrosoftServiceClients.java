package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;

import java.util.concurrent.TimeUnit;

/**
 * Channels and stubs to this repo's {@code MicrosoftService}. One channel
 * per GCA request: the admin-supplied target can differ per connection.
 */
public final class MicrosoftServiceClients {

    private MicrosoftServiceClients() {
    }

    /**
     * Opens a channel to {@code config}'s MicrosoftService target.
     *
     * @param config target and plaintext flag
     * @return a new channel; caller must {@link #shutdown(ManagedChannel)}
     */
    public static ManagedChannel channel(ConnectorCustomConfig config) {
        ManagedChannelBuilder<?> builder = ManagedChannelBuilder.forTarget(config.target());
        if (config.plaintext()) {
            builder.usePlaintext();
        }
        return builder.build();
    }

    /**
     * A blocking stub on {@code channel}.
     *
     * @param channel an open MicrosoftService channel
     * @return a blocking stub
     */
    public static MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub(ManagedChannel channel) {
        return MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    /**
     * Shuts {@code channel} down, interrupting if it does not terminate promptly.
     *
     * @param channel the channel to close; {@code null} is ignored
     */
    public static void shutdown(ManagedChannel channel) {
        if (channel == null) {
            return;
        }
        channel.shutdown();
        try {
            if (!channel.awaitTermination(5, TimeUnit.SECONDS)) {
                channel.shutdownNow();
            }
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            channel.shutdownNow();
        }
    }
}
