package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.MethodDescriptor;
import io.grpc.ServiceDescriptor;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Checks that {@code docs/permissions.md} names every rpc. */
class PermissionsDocTest {

    /** The record of what each rpc needs, relative to this module. */
    private static final Path DOC = Path.of("..", "docs", "permissions.md");

    @Test
    void everyRpcIsNamedInThePermissionsDoc() throws IOException {
        String doc = Files.readString(DOC, StandardCharsets.UTF_8);
        List<String> missing = new ArrayList<>();
        for (ServiceDescriptor service : List.of(MicrosoftServiceGrpc.getServiceDescriptor(),
                MicrosoftChangeServiceGrpc.getServiceDescriptor(),
                MicrosoftAdminServiceGrpc.getServiceDescriptor())) {
            for (MethodDescriptor<?, ?> method : service.getMethods()) {
                String rpc = method.getBareMethodName();
                if (!doc.contains("`" + rpc + "`")) {
                    missing.add(rpc);
                }
            }
        }
        assertThat(missing)
                .as("rpcs with no row in docs/permissions.md: say what each needs and how that is known")
                .isEmpty();
    }
}
