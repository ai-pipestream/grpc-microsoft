package ai.pipestream.microsoft;

import java.io.IOException;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

/** Sends the HTTP requests of a {@link GraphClient}. */
@FunctionalInterface
interface GraphTransport {

    /**
     * Sends one request and waits for its response.
     *
     * @param request the request
     * @param handler reads the response body
     * @param <T> the body type
     * @return the response
     * @throws IOException if the exchange fails
     * @throws InterruptedException if the call is interrupted
     */
    <T> HttpResponse<T> send(HttpRequest request, HttpResponse.BodyHandler<T> handler)
            throws IOException, InterruptedException;
}
