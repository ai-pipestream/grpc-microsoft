# Contributing

This is the Microsoft Graph connector of the ai.pipestream gRPC family.
Patches should look like they belong in an Apache-style Java project:
typed wire, comments that pass lint, tests that do not need the internet.

Build with a grpc-integration checkout next to this one (see the README);
shared modules come from there.

## Before you send a PR

```
./gradlew build          # unit/fake tests + javadoc (-Xdoclint:all -Xwerror)
buf lint                 # STANDARD + COMMENTS on grpc-microsoft-api
```

`buf` is not a Gradle plugin here (codegen stays protoc). Install from
https://buf.build or use the version CI pins (`1.54.0`). Microsoft's
GCA `Contracts/` tree is a frozen copy and is not in the Buf workspace.

## Rules that already bit us

- **No JSON bridge** on the typed gRPC path. Domain messages stay protobuf.
- **Shared code goes to grpc-integration.** Anything source-neutral (the
  ledger, connection catalog, output stores, OKF writer, MCP server,
  Kafka Connect base classes) belongs there, not here. This repo plugs in
  through `OkfCatalogSource`, `OutputFormat`, `McpToolProvider`, and Kafka
  Connect's `SourceConnector` ServiceLoader entries. A sink that talks to
  `SyncTableService` lives in the caller and depends on `grpc-sync-api`
  only.
- **Wire compatibility.** Proto packages and Java packages do not change.
  `buf breaking` with `WIRE_JSON` against `main` should stay clean.
- **Secrets stay out of git, logs, and PR text.** Live smokes read env /
  GitHub Actions secrets; they never print tokens.
- **Buf COMMENTS is strict.** Every service, rpc, message, field, oneof,
  enum, and enum value needs a non-empty comment. STANDARD applies too
  (`RPC_REQUEST_RESPONSE_UNIQUE`, enum prefixes, ...).
- **Javadoc is strict** on handwritten public/protected API
  (`-Xdoclint:all -Xwerror`). Generated protobuf is excluded.
- **Virtual threads.** gRPC handlers and MCP tool calls are blocking by
  design; the servers run them on virtual threads. Do not introduce a
  platform-thread pool for I/O.
- **Live tests are read-only.** No full `Sync` crawl in CI. The smoke is
  client-credentials `ListSites` / `ListDrives` — not `/me`.

## Layout

See the grpc-integration `docs/architecture.md` for the process topology.
`*-api` is proto, `*-service` is the process, `*-mcp` and `*-connect` are
the MCP and Kafka Connect plug-ins.
