/**
 * Kafka Connect source plugin that emits raw MicrosoftChange protobuf
 * bytes, built on the grpc-integration {@code grpc-connect} base classes.
 * It can crawl through the in-process Graph client or pull a {@code Sync}
 * stream from MicrosoftService.
 */
package ai.pipestream.connect;
