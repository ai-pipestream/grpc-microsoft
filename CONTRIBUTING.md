# Contributing

This is a small, MIT-licensed Java gRPC repo. Patches should look like they
belong in an Apache-style Java project: typed wire, comments that pass
lint, tests that do not need the internet.

## Before you send a PR

```
./gradlew build          # unit/fake tests + javadoc (-Xdoclint:all -Werror)
buf lint                 # STANDARD + COMMENTS on our protos
```

`buf` is not a Gradle plugin here (codegen stays protoc). Install from
https://buf.build or use the version CI pins (`1.54.0`). Microsoft's
GCA `Contracts/` tree is a frozen copy and is not in the Buf workspace.

## Rules that already bit us

- **No JSON bridge** on the typed gRPC path. Domain messages stay protobuf.
- **Do not invent shared modules.** A sink that talks to `SyncTableService`
  lives in the caller and depends on `grpc-sync-api` only. `grpc-okf` is the
  dedicated OKF/WARC library both crawlers share; do not fold that producer
  into either service. The output SPI (`grpc-output-spi` plus filesystem /
  S3 ServiceLoader jars) is the dedicated destination seam; do not fold
  S3 into either service.
- **Secrets stay out of git, logs, and PR text.** Live smokes read env /
  GitHub Actions secrets; they never print tokens.
- **Buf COMMENTS is strict.** Every service, rpc, message, field, oneof,
  enum, and enum value needs a non-empty comment. STANDARD applies too
  (`RPC_REQUEST_RESPONSE_UNIQUE`, enum prefixes, ...).
- **Javadoc is strict** on handwritten public/protected API
  (`-Xdoclint:all -Werror`). Generated protobuf is excluded.
- **Virtual threads.** gRPC handlers and MCP tool calls are blocking by
  design; the servers run them on virtual threads. Do not introduce a
  platform-thread pool for I/O.
- **Live tests are read-only.** No full `Sync` crawl in CI. Confluence
  smokes `ListSpaces` / homepage / attachments. Microsoft smokes
  client-credentials `ListSites` / `ListDrives` — not `/me`.

## Layout

See [docs/architecture.md](docs/architecture.md). Module naming follows
the gRPOIc split: `*-api` is proto, `*-service` is the process.
