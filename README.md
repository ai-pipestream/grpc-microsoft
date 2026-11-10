# grpc-microsoft

Microsoft Graph as typed gRPC. The proxy speaks Graph REST upstream and
gRPC downstream for SharePoint sites, drives, and files, with a
cursor-based `Sync` stream, an asset ledger, OKF / WARC output (optionally
uploaded back to a SharePoint library), a Kafka Connect source, MCP tools,
and a Microsoft Copilot (Graph Connector Agent) connector process.

The shared, source-neutral pieces (sync ledger and connection catalog,
output SPI and stores, OKF writer, the combined MCP server, Kafka Connect
base classes) live in
[grpc-integration](https://github.com/ai-pipestream/grpc-integration).
This repo was split out of
[grpc-confluence](https://github.com/ai-pipestream/grpc-confluence) with
its history.

This repository is licensed under Apache-2.0. The third-party Microsoft
contracts retain their MIT license and notices.

The Microsoft Copilot connector wire contracts are copied from
[Custom-Copilot-Connector-using-Connector-SDK](https://github.com/microsoft/Custom-Copilot-Connector-using-Connector-SDK)
(MIT). See `grpc-microsoft-connector/NOTICE`.

## Modules

Group `ai.pipestream.microsoft`, version `0.1.0-SNAPSHOT`.

| Module | Port | What it is |
|---|---|---|
| `grpc-microsoft-api` | | The API proto contract (`ai.pipestream.microsoft.v1`). 1:1 Microsoft Graph surface we crawl, validated for gRPC. |
| `grpc-microsoft-service` | 9096 | Live Graph proxy for sites, drives, and files. |
| `grpc-microsoft-output` | | Graph store keys, OKF catalog mapping, the `microsoft-connector` output format, and its `OkfCatalogSource` registration. |
| `grpc-microsoft-connector` | 30303 | Microsoft Copilot / GCA adapter that forwards crawls to MicrosoftService. |
| `grpc-microsoft-mcp` | | `microsoft_*` tools for the combined `grpc-mcp` server (`McpToolProvider`). `run` starts that server with these tools. |
| `grpc-microsoft-connect` | | Kafka Connect source (`MicrosoftSourceConnector`) and its plugin zip. |

## Building

The core modules are Maven dependencies
(`ai.pipestream.integration:grpc-sync-api:0.1.0-SNAPSHOT` and friends).
There are two ways to get them:

1. **Sibling checkout (recommended).** Clone grpc-integration next to
   this repo. `settings.gradle.kts` sees `../grpc-integration` and
   includes it as a Gradle composite build, so the core builds from source
   and no registry access is needed:

   ```
   git clone https://github.com/ai-pipestream/grpc-integration
   git clone https://github.com/ai-pipestream/grpc-microsoft
   cd grpc-microsoft
   ./gradlew build
   ```

   A checkout elsewhere: `./gradlew build -PintegrationDir=/path/to/grpc-integration`.

2. **Published snapshots.** Without a checkout the snapshots resolve from
   the ai.pipestream Forgejo registry
   (`https://git.rokkon.com/api/packages/ai-pipestream/maven`, no
   credentials needed to read).

CI checks out both repositories side by side and uses option 1.

## Microsoft Graph proxy

```
MICROSOFT_TENANT_ID=...
MICROSOFT_CLIENT_ID=...
MICROSOFT_CLIENT_SECRET=...
MICROSOFT_SITE_ID=...               # optional; empty = signed-in user's drive
MICROSOFT_DRIVE_IDS=...             # optional comma-separated allowlist
MICROSOFT_GRPC_PORT=9096
```

```
./gradlew :grpc-microsoft-service:run
```

RPCs: `GetMe`, `ListSites`, `ListDrives`, `ListChildren` (stream), `GetItem`,
`DownloadItem`, `Sync` (stream). Handlers run on virtual threads.

Point the proxy at a running sync-table (from grpc-integration) to record
crawl / update / delete rows:

```
SYNC_TABLE_TARGET=localhost:9097
SYNC_TABLE_PLAINTEXT=true
```

Completed `Sync` runs write through the grpc-integration output SPI
(`OUTPUT_DIR` / `OKF_DIR`, `OUTPUT_STORE`, `OUTPUT_FORMATS`, ...; see the
grpc-integration README). This repo adds the `microsoft-connector` format
and the Graph OKF catalog mapping.

`GetItem`, `DownloadItem`, and `Sync` flatten SharePoint list-item columns
into typed `ListColumn`s (scalars and one-level nested objects; `@odata.*`
skipped). `ListChildren` stays a listing RPC and does not fetch columns.

To land the OKF payload **on** a SharePoint library (markdown tree + zip +
`.warc.gz`), set a destination drive. Files larger than 4 MiB use a Graph
upload session; the session URL is called **without** `Authorization`.
Needs `Files.ReadWrite.All` (or equivalent) on the app. Live CI smokes stay
read-only and do not set these.

```
OKF_SPO_DRIVE_ID=b!...
OKF_SPO_FOLDER_PATH=/Knowledge/okf-run
```

## Copilot connector (GCA)

A separate process that implements Microsoft's four connector services
(`ConnectorInfoService`, `ConnectionManagementService`,
`ConnectorCrawlerService`, `ConnectorOAuthService`) and forwards crawls to
`MicrosoftService`. Graph tokens stay on the Microsoft proxy.

Connector id (stable, put this in the GCA manifest):
`43760992-c66e-46bd-b937-b63e021aa63b`

Custom configuration JSON (GCA does not interpret this string):

```json
{
  "target": "localhost:9096",
  "plaintext": true,
  "driveIds": [],
  "folderPath": "/",
  "includeContent": false
}
```

`ValidateAuthentication` calls `GetMe`. When custom configuration is empty,
a `host:port` or `grpc://host:port` datasource URL is treated as the
MicrosoftService target.

```
MICROSOFT_GRPC_TARGET=localhost:9096
CONNECTOR_GRPC_PORT=30303
./gradlew :grpc-microsoft-connector:run
```

## MCP tools

`grpc-microsoft-mcp` registers `MicrosoftMcpToolProvider`
(`META-INF/services/ai.pipestream.mcp.McpToolProvider`). Put its jar, with
`grpc-microsoft-api`, on the classpath of the combined `grpc-mcp` server
from grpc-integration (for the Docker image: `/opt/mcp/ext`) and it adds:

```
microsoft_get_me / microsoft_sync
```

The server dials `MICROSOFT_GRPC_TARGET` (default `localhost:9096`). To
run that server with the Microsoft tools from this repo (plus the core
`app_*`, `connection_*`, and `sync_table_*` tools):

```
MCP_PORT=8090
MICROSOFT_GRPC_TARGET=localhost:9096
SYNC_TABLE_TARGET=localhost:9097
./gradlew :grpc-microsoft-mcp:run
```

## Kafka Connect

`./gradlew :grpc-microsoft-connect:connectPluginZip` builds
`grpc-microsoft-connect/build/distributions/grpc-microsoft-connect-0.1.0-SNAPSHOT-plugin.zip`.
Drop it on a Connect worker `plugin.path`. Values are raw `MicrosoftChange`
protobuf bytes. Prefer `grpc.target` so the worker talks to the
already-running proxy (one credential surface).

```json
{
  "name": "microsoft-source",
  "config": {
    "connector.class": "ai.pipestream.connect.MicrosoftSourceConnector",
    "topic": "microsoft.changes",
    "grpc.target": "localhost:9096",
    "grpc.plaintext": "true"
  }
}
```

## Proto lint (Buf)

`grpc-microsoft-api` lints with Buf **STANDARD** plus **COMMENTS**.
Microsoft's GCA `Contracts/` tree is a frozen MIT copy of their wire and is
**not** in the Buf workspace. Its proto package stays
`Microsoft.Graph.Connectors.Contracts.Grpc`.

```
buf lint
```

## Build and tests

Java 25 toolchain. `./gradlew build` is the fake/unit suite plus javadoc
(`-Xdoclint:all -Xwerror`) and never talks to Graph. Every compiler
warning and Gradle deprecation fails the build.

The live smoke is **read-only**, excluded from `test`, and skips unless
credentials are in the environment:

```
./gradlew :grpc-microsoft-service:liveSmokeTest
```

It uses **Entra application** (client credentials), not a user mailbox
login. `/me` does not exist on an app-only token, so the probe is
`ListSites` and, when a site id is known, `ListDrives`. Create an app
registration with application permissions `Sites.Read.All` and
`Files.Read.All` (admin consent), then a client secret.

### GitHub Actions secrets

After they are set, `live-microsoft` runs on same-repo PRs and pushes
(fork PRs never receive secrets). Missing secrets skip the step.

| Secret | Required |
|---|---|
| `MICROSOFT_TENANT_ID` | yes |
| `MICROSOFT_CLIENT_ID` | yes |
| `MICROSOFT_CLIENT_SECRET` | yes |
| `MICROSOFT_SITE_ID` | no (first `ListSites` hit is used for drives) |
| `MICROSOFT_DRIVE_IDS` | no |

Do not paste tokens into issues, PR text, or agent chat.

## Docker

```
docker build -t grpc-microsoft -f Dockerfile.microsoft .
docker build -t grpc-microsoft-connector -f Dockerfile.connector .
# build against a local core checkout instead of the published snapshots:
docker build --build-context integration=../grpc-integration -t grpc-microsoft -f Dockerfile.microsoft .
```
