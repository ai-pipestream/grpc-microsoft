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
| `grpc-microsoft-admin` | 9097 | `MicrosoftAdminService`: lists, ensures and revokes the per-site grants of applications that hold `Sites.Selected`. The only part that runs with `Sites.FullControl.All`; deployed on its own. |

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

Write RPCs (need a `Sites.Selected` site grant of `write` or higher, see
[Permissions](#permissions)): `UploadItem` (bidi stream; upload sessions
for large files, resumable), `CreateFolder`, `UpdateDriveItem`,
`DeleteDriveItem`, `RestoreDriveItemVersion`, `CopyDriveItem` (stream),
`CreateListItem`, `UpdateListItemFields`, `DeleteListItem`,
`EnsureColumns`, `EnsureList`, and `Mutate` (bidi stream; packs writes
into Graph `$batch` requests). Updates and deletes take `expected_etag`
(sent as `If-Match`); a mismatch is `ABORTED`.

`Sync` with `include_security` also reads SharePoint's own security: each
site's groups with their members (`site_group`), the role assignments of the
site and of every list and item that has its own (`site_acl`), and on every
drive, list and item the id of the `site_acl` that applies to it
(`acl_owner_id`). A pass resumed from a cursor also reads the site's change
log: a list item that stopped inheriting, or inherits again, is emitted
again with everything below it even when no delta reports it, and role
assignments that are gone are emitted as deletes. `SetSiteRoleInheritance` makes a list or list item stop
inheriting, or inherit again. `CreateSiteGroup` and `DeleteSiteGroup`,
`AddSiteGroupMember` and `RemoveSiteGroupMember` change SharePoint groups;
`AddSiteRoleAssignment` and `RemoveSiteRoleAssignment` grant or take one
permission level (`ListSiteRoleDefinitions` lists them) on a site, list or
list item that has its own role assignments. All need the certificate
credential and the site grant `fullcontrol`.

Point the proxy at a running sync-table (from grpc-integration) to record
crawl / update / delete rows:

```
SYNC_TABLE_TARGET=localhost:9097
SYNC_TABLE_PLAINTEXT=true
```

A sync-table write that fails ends the `Sync` with an error whose text names
the ledger RPC and the row: `UNAVAILABLE` when the ledger could not be
reached in time (run the `Sync` again from the cursor you hold), `INTERNAL`
for any other refusal. No resume cursor is sent, so the delta link never
moves past a change the ledger did not record, and the ledger is not
reconciled against a run whose rows did not all land.

The ledger is reconciled (rows the run did not write are marked deleted)
only after a complete run: no `since_cursor`, at least one drive or list
swept, and no narrower scope in the request than the connection's own. A
`Sync` that names some of the connection's drives or sites, or a folder,
never looked at the rest, so it leaves their rows and the output mirror as
they were. Naming exactly the configured drives, or the configured site,
still counts as complete.

A complete run reconciles one kind of entity at a time, and only the kinds
it read in full: sites, drives and drive items always; lists and list items
only with `include_lists`; site groups and site admins only with
`include_security`; ACLs only with `include_security` and `include_lists`
together, because a site's lists hold ACLs too (a connection with a drive
allowlist and no site has no lists, so there `include_security` is enough).
A full `Sync` that leaves a flag out never read those rows and leaves them,
and their records in the output mirror, as an earlier `Sync` wrote them.

A run that is not complete still reconciles each drive and list it read
from scratch on its own. That is the drive whose delta link Graph rejected
with `410 Gone`: it is enumerated again without delete records, so the
items gone from it are marked deleted in the ledger and leave the output
mirror, while the drives still on their delta links are not touched. The
same holds for a list, and for a `Sync` of some of the connection's drives
without a `since_cursor`. A drive read under a `folder_path` other than the
connection's own is not reconciled. Ledger rows carry the drive id, or
`site_id/list_id`, as `Asset.scope`, and the run calls the sync table's
`ReconcileScope` rpc: a sync table too old to have it answers
`UNIMPLEMENTED`, which fails the `Sync` and reconciles nothing. Rows
written before rows carried a scope get one the next time they are
written; until then only a complete run reconciles them.

The output and OKF bundle is one for the process, shared by every
connection. Each record in it belongs to the connection whose `Sync` wrote
it: a complete `Sync` of one connection replaces that connection's records
and leaves the others in the bundle, and two `Sync` calls running at once
each write their own records when they finish. A connection with no
complete `Sync` yet in this process writes an incremental pass beside the
bundle, under `increments/<connection id>/<run id>`.

Completed `Sync` runs write through the grpc-integration output SPI
(`OUTPUT_DIR` / `OKF_DIR`, `OUTPUT_STORE`, `OUTPUT_FORMATS`, ...; see the
grpc-integration README). This repo adds the `microsoft-connector` format
and the Graph OKF catalog mapping.

In the OKF / WARC output a `resource` record exists only for bytes read
from Graph. `Sync` without `include_content`, a file over
`max_inline_bytes`, and every folder arrive as metadata. Those items keep
their OKF markdown, marked `capture: not_fetched` in its frontmatter, and
their link on the collection page, and they get no `resource` record:
nothing is written in place of content that was not fetched. A drive item
is named by its `web_url` (a `urn:okf:0.2:microsoft:item:<drive>/<item>`
when Graph sent none), never by `download_url`: that link is
pre-authenticated and short-lived, so it is not written to the bundle.

`GetItem`, `DownloadItem`, and `Sync` flatten SharePoint list-item columns
into typed `ListColumn`s (scalars and one-level nested objects; `@odata.*`
skipped). `ListChildren` stays a listing RPC and does not fetch columns.

A `Sync` with `include_content` sets `DriveItem.content_state` on every
drive item: `INLINED` (the bytes are in `content`), `EMPTY`,
`OMITTED_OVER_LIMIT` (larger than `max_inline_bytes` or the server's
`MICROSOFT_ATTACHMENT_MAX_BYTES`, by its listed size or by what the
download turned out to be; read it with `DownloadContent`),
`OMITTED_NOT_A_FILE` (a folder or a package) and `OMITTED_MALWARE`. An
item without `content` is never ambiguous. A read that did not ask for
content leaves the state `UNSPECIFIED`.

Reads ask Graph for an explicit property list, because `sharepointIds`,
`publication`, the `system` facet and the download URL only come when
selected. When Graph rejects that query (a 400 naming `$select`,
`$expand`, a query option or an unknown property) the read is answered
from Graph's default properties, and the entities lack those fields. That
is never silent: it is logged at WARNING, and the rpc ends with a typed
`DegradedProjections` message in its trailing metadata (key
`ai.pipestream.microsoft.v1.DegradedProjections-bin`) naming the route, the
rejected query, Graph's message and how many reads were answered that way.
A call whose reads were all answered in full carries no such trailer. Any
other 400 fails the rpc; it is not retried without the query.

To land the OKF payload **on** a SharePoint library (markdown tree + zip +
`.warc.gz`), set a destination drive. Files larger than 4 MiB use a Graph
upload session; the session URL is called **without** `Authorization`.
Needs `Files.ReadWrite.All` (or equivalent) on the app. Live CI smokes stay
read-only and do not set these.

```
OKF_SPO_DRIVE_ID=b!...
OKF_SPO_FOLDER_PATH=/Knowledge/okf-run
```

## Change feed

`MicrosoftChangeService` (same port) streams changes continuously.
`WatchChanges` runs `Sync` passes back to back and polls by default. Graph
change notifications only shorten the wait, and need a receiver Graph can
reach over HTTPS:

| Variable | What |
|---|---|
| `MICROSOFT_NOTIFICATION_PORT` | Port of the notification receiver (`/graph/notifications`, `/graph/lifecycle`); unset = receiver off |
| `MICROSOFT_NOTIFICATION_CLIENT_STATE` | Secret sent with every subscription and checked on every notification; required for the receiver and for `CreateSubscription` |
| `MICROSOFT_SUBSCRIPTION_RENEWAL` | `off` stops the server renewing its subscriptions; default on when a clientState is set |
| `MICROSOFT_SUBSCRIPTION_RENEW_BEFORE_MINUTES` | How long before its expiration a subscription is renewed; default 1440 |

The server keeps its own subscriptions alive: the ones created through
`CreateSubscription` and, after a restart, the ones Graph lists with its
clientState (default connection at start, other connections when
`ListSubscriptions` or `GetSubscription` is called for them). Graph leaves the
clientState out of a listing, so each listed subscription the server does not
keep yet is read once more on its own. It renews each
to the longest expiration Graph allows before it expires and when Graph asks
for a reauthorization, and creates a subscription again when Graph no longer
has it. `ListSubscriptionRenewals` shows what it keeps and the last error. A
subscription that expires while the server is down is not created again.

`WatchChangesRequest.directory` adds Entra users and groups to every pass:
`user_change` and `group_change` messages (a group change carries its member
changes when `group_members` is set), then one `directory_snapshot` per
kind. The first pass lists every user and group; later passes read
`/users/delta` and `/groups/delta` from the links the resume cursor carries.
Subscriptions on `/users`, `/groups` and `/groups/{id}/members` shorten the
wait the same way as the ones on drives and lists.

`WatchChangesRequest.site_changes` adds the SharePoint change log of the named
sites to every pass: one `site_change` per entry, then one
`site_change_snapshot` per site. The log names what a drive notification with
security webhooks only hints at: the list item, list or site whose role
assignments or inheritance changed, and the group whose members changed. The
first pass of a site takes only the end of its log; the resume cursor carries
the change token of each site. Changes to SharePoint groups and their members
are only in the site collection's log (`scope`), and name the group, not the
member. `ListSiteChanges` reads the same log on its own. Both need the
certificate credential.

## Permissions

[docs/permissions.md](docs/permissions.md) is the record of the least
privilege each rpc needs, and of which rows are proven against a real tenant.
In short: the connector app holds Graph and SharePoint Online
`Sites.Selected` and one grant per site (`read` to index, `manage` to write,
`fullcontrol` to read or change access control), plus Graph `User.Read.All`
and `GroupMember.Read.All` for Entra users and groups; rpcs that call
SharePoint REST need the certificate credential. Only the admin service
holds `Sites.FullControl.All`.

## Site grants (admin service)

The connector app holds `Sites.Selected` and sees only the sites it was
granted. `MicrosoftAdminService` (module `grpc-microsoft-admin`) hands out
those grants and is the only process that signs in with an identity holding
`Sites.FullControl.All`. Run it only when grants need to change.

| rpc | What it does | Graph |
|---|---|---|
| `ListSiteAppGrants` | Applications granted on a site, with their roles | `GET /sites/{id}/permissions`, `GET /sites/{id}/permissions/{id}` |
| `EnsureSiteAppGrant` | Gives an application exactly one role on a site (create, change, or leave alone); `dry_run` reports what would change | `POST` / `PATCH /sites/{id}/permissions` |
| `RevokeSiteAppGrant` | Removes an application's grants on a site | `DELETE /sites/{id}/permissions/{id}` |

| Variable | Required |
|---|---|
| `MICROSOFT_TENANT_ID` | yes |
| `MICROSOFT_ADMIN_CLIENT_ID` | yes |
| `MICROSOFT_ADMIN_CLIENT_CERT_PEM` | yes, unless the secret is set |
| `MICROSOFT_ADMIN_CLIENT_SECRET` | yes, unless the certificate is set |
| `MICROSOFT_ADMIN_GRPC_PORT` | no (default 9097) |

```
./gradlew :grpc-microsoft-admin:run
```

`./gradlew :grpc-microsoft-admin:liveAdminTest` runs the admin scenarios
against a real tenant (also needs `MICROSOFT_SITE_ID` and `MICROSOFT_CLIENT_ID`)
and records the scrubbed Graph exchanges under
`grpc-microsoft-admin/src/test/resources/recordings/admin`. The default `test`
task replays them with no network. Reading grants and the dry runs change
nothing; the create, change and revoke scenario runs only with
`MICROSOFT_ADMIN_LIVE_CHANGE=1`.

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

Crawl checkpoints: `customMarkerData` is only ever a `Sync` resume cursor,
or empty. Every checkpoint written while a crawl runs repeats the cursor
the request carried, and the last one carries the cursor the pass ended
with. An incremental crawl hands the stored cursor back as
`SyncRequest.since_cursor`, so it is a delta pass: only what changed since
that cursor, with each delete as a `DeletedItem`. With no stored cursor it
is a full enumeration, which reports no deletes. A full crawl never sends a
cursor. A `Sync` stream that ends without its terminal resume cursor was
cut short: the crawl ends with `DatasourceError` and no new checkpoint.

Item ids: an item is indexed as `<drive id>/<item id>`, and a `DeletedItem`
is sent under that same id, read from the drive item the delete carries. A
delete that does not carry its drive item cannot name what it removes: the
crawl ends with `DatasourceError` and no new checkpoint, since a delete
sent under any other id would be acknowledged and remove nothing.

Content: with `includeContent` on, a file's bytes become the item's content
only when they are text. That is `text/*` (not RTF), JSON, XML, YAML,
JavaScript and the `+json` / `+xml` types, with HTML and XHTML sent as
HTML. Office documents, PDFs and images are not text: their items carry
properties only until a parser service turns them into text. The bytes are
decoded strictly, by the byte order mark, then the media type's `charset`,
then the charset an HTML or XML document declares, then UTF-8. The crawl
asks `Sync` to inline at most 3 MiB per file, and sends at most 3 MiB of
UTF-8 text per item. An item whose content is not sent is still emitted,
and its crawl bit reports `PartialSuccess` with a message that names the
item, the reason (`NOT_TEXT`, `TOO_LARGE`, `NOT_INLINED`, `NOT_DECODABLE`)
and what was found. The last bit of the crawl counts those items. With
`includeContent` off, nothing is reported: no content was asked for.

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

The Connect offset is the pass's resume cursor (an encoded `SyncCursor`),
in both modes. Each poll resumes from it, so after the first full
enumeration a poll publishes only what changed, deletes included
(`CHANGE_OPERATION_DELETE`). A `Sync` stream that ends without its resume
cursor fails the poll, and nothing of it is committed. An offset written by
a version that stored the poll's wall-clock time is read as "no cursor":
the next pass is one full enumeration, logged at WARNING.

Direct mode (no `grpc.target`) takes `microsoft.tenant.id`,
`microsoft.client.id`, `microsoft.client.secret`, and optionally
`microsoft.site.id`, `microsoft.drive.ids`, `microsoft.folder.path`,
`microsoft.graph.base.url` and `microsoft.authority` (the Entra authority
host, for a national cloud whose Graph root the base URL names).

`include.content=true` inlines file bytes on each drive item, in both
modes. `max.inline.bytes` caps the file size (0 = no limit in direct mode,
the service's own limit in gRPC mode); keep it under the producer's
`max.request.size`, or one large file fails the task. A record whose file
was not inlined says why in `DriveItem.content_state`.

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
`ListSites` and, when a site id is known, `ListDrives`. The app registration
and what it must hold are in [docs/permissions.md](docs/permissions.md).

The live write test is opt-in. It calls every write RPC inside one new
folder (`pipestream-live-write-<timestamp>-<random>` at the drive root)
and one new list of the same name, removes both at the end, writes
`build/live-write.md`, and records the scrubbed Graph exchanges that the
default tests replay (see
`grpc-microsoft-service/src/test/resources/recordings/README.md`). It needs write permission on the site and skips
unless `MICROSOFT_SITE_ID` and `MICROSOFT_LIVE_WRITE=1` are set next to
the credentials:

```
MICROSOFT_LIVE_WRITE=1 ./gradlew :grpc-microsoft-service:liveWriteTest
```

### GitHub Actions secrets

After they are set, `live-microsoft` runs on same-repo PRs and pushes
(fork PRs never receive secrets). Missing secrets skip the step.

| Secret | Required |
|---|---|
| `MICROSOFT_TENANT_ID` | yes |
| `MICROSOFT_CLIENT_ID` | yes |
| `MICROSOFT_CLIENT_SECRET` | yes, unless the certificate is set |
| `MICROSOFT_CLIENT_CERT_PEM` | no (private key and certificate PEM; used instead of the secret when set) |
| `MICROSOFT_SITE_ID` | no (first `ListSites` hit is used for drives) |
| `MICROSOFT_DRIVE_IDS` | no |

A catalog connection (`connection_id`) and an inline `ProbeConnection` carry
the same choice: `client_secret`, or `client_certificate_pem` to sign in with
a certificate. SharePoint site security reads need the certificate.

Do not paste tokens into issues, PR text, or agent chat.

## Docker

```
docker build -t grpc-microsoft -f Dockerfile.microsoft .
docker build -t grpc-microsoft-connector -f Dockerfile.connector .
docker build -t grpc-microsoft-admin -f Dockerfile.admin .
# build against a local core checkout instead of the published snapshots:
docker build --build-context integration=../grpc-integration -t grpc-microsoft -f Dockerfile.microsoft .
```
