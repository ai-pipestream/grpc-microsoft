# AGENTS.md: grpc-microsoft

Read the workspace guide first:
[grpc-integration/AGENTS.md](https://github.com/ai-pipestream/grpc-integration/blob/main/AGENTS.md).
It has the repository map, where the open work lives, the standing rules,
the testing model and what to focus a review on. This file adds what is
specific to the Microsoft (SharePoint Online, OneDrive, Entra) connector.

## Which branch this describes

`main`: the reads, the write rpcs, the Entra and SharePoint security reads,
term columns, the admin service and the change stream
(`MicrosoftChangeService`).

## Build

Check out [grpc-integration](https://github.com/ai-pipestream/grpc-integration)
next to this repository (same branch name when it exists), then:

```
./gradlew build      # unit, fake-server and replay tests; javadoc; never calls Microsoft
buf lint             # STANDARD + COMMENTS on grpc-microsoft-api, buf 1.54.0
```

Java 25. Any compiler, javadoc or Gradle deprecation warning fails the
build. Do not change how tests are run.

## Modules

| Module | What |
|---|---|
| `grpc-microsoft-api` | `ai.pipestream.microsoft.v1`. `microsoft_service.proto` is the service; `drive.proto`, `list.proto`, `site.proto`, `permission.proto`, `write.proto`, `directory.proto` (Entra), `site_security.proto` (SharePoint groups, role assignments, `SiteAcl`, `SiteAdmins`), `site_change.proto` (the SharePoint change log), `term_store.proto`, `microsoft_admin_service.proto` and `change_feed.proto` (`MicrosoftChangeService`). |
| `grpc-microsoft-service` | The process (port 9096). `MicrosoftGrpcService` (rpc handlers), `GraphClient`, `GraphAuth` and `ClientCertificate` (secret or certificate sign-in), `MicrosoftSessions` (catalog connections), `MicrosoftMapper`, `MicrosoftCrawler` (`Sync`, delta), `WritePlans`, `UploadHandler`, `MutateHandler` (`$batch`), `ListFieldWriter` and `ColumnSpecs` (typed column values), `SharePointRest`, `SiteSecurityReader`, `SiteSecurityWriter`, `SiteSecurityMapper` and `SiteChanges` (`getchanges`), `DirectoryWriter` (Entra group writes). Change stream: `MicrosoftChangeGrpcService`, `GraphSubscriptions`, `SubscriptionRenewer`, `NotificationReceiver`, `NotificationBus`. |
| `grpc-microsoft-admin` | `MicrosoftAdminService`: list, ensure (with `dry_run`) and revoke per-site application grants. The only process that signs in with `Sites.FullControl.All`; deployed on its own. Default port 9097, the same as the sync ledger's default. |
| `grpc-microsoft-connector` | Microsoft Copilot / Graph Connector Agent adapter (port 30303). Its `Contracts/` protos are a frozen MIT copy of Microsoft's wire and are outside the Buf workspace. |
| `grpc-microsoft-output` | Store keys, OKF catalog mapping, the `microsoft-connector` output format. |
| `grpc-microsoft-mcp` | `microsoft_*` tools for the combined MCP server. |
| `grpc-microsoft-connect` | Kafka Connect source. |

## Contract references

- `docs/contracts/microsoft-gap-analysis.md`: entities and rpcs against
  Graph v1.0, design decisions, what is left out and why.
- Graph first. SharePoint REST is used only where Graph has no route or
  refuses the operation: site groups and members, role assignments and
  inheritance, site admins, the change log (`getchanges`), hyperlink and term column creation
  (`CreateFieldAsXml`), hyperlink values (`ValidateUpdateListItem`).
  SharePoint REST app-only calls need the certificate credential.
- The README's "Permissions" table lists which permission and site grant
  each group of rpcs needs, as proven on the test tenant.

## Tests

| Suite | Live task (opt-in) | Replay in `test` | Recordings |
|---|---|---|---|
| `WriteScenarios` (writes and security reads) | `liveWriteTest` (`MICROSOFT_LIVE_WRITE=1`, `MICROSOFT_SITE_ID`) | `MicrosoftWriteReplayTest` | `grpc-microsoft-service/.../recordings/microsoft/` |
| `ChangeScenarios` (change stream, subscription renewal, Entra users and groups) | `liveChangesTest` | `MicrosoftChangeReplayTest` | `.../recordings/microsoft-changes/` |
| `AdminGrantScenarios` | `:grpc-microsoft-admin:liveAdminTest` (grant changes only with `MICROSOFT_ADMIN_LIVE_CHANGE=1`) | `MicrosoftAdminReplayTest` | `grpc-microsoft-admin/.../recordings/admin/` |
| read coverage | `liveCoverageTest` | none | report only |
| smoke | `liveSmokeTest` | none | none |

Recording and scrubbing rules:
`grpc-microsoft-service/src/test/resources/recordings/README.md`.
`RecordingScrubber` and the replay transports are in the service module's
`testFixtures` so the admin module shares them. `RecordingsHygieneTest`
checks every committed file in every build, including percent-encoded
email addresses.

Live runs go through `.github/workflows/ci.yml` `workflow_dispatch` inputs
(`run_live_write`, `run_live_coverage`, `run_live_admin`,
`import_recordings_run_id`, `import_admin_recordings_run_id`,
`run_live_changes`). Some inputs change tenant grants.
Reviewers do not dispatch them.

## Microsoft specifics to keep in mind

- Live behaviour found so far (each has a test or a mapping): Graph answers
  411 to a POST with no body, so bodyless actions send `{}`; Graph refuses
  hyperlink and term column creation and hyperlink value writes; SharePoint
  drives silently drop `driveItem.description`; Graph's term store does not
  serve application tokens, so term store reads are unproven; group members
  and role assignments need the site grant `fullcontrol`, not `manage`;
  upload sessions are called without `Authorization`; the SharePoint
  change log shows an inheritance break as `SCOPE_ADD` and a reset as
  `SCOPE_DELETE` on the item, and accepts a token older than the log;
  group and group member changes are only in the site collection log and
  name the group, not the member; Graph asks for a reauthorization of a
  50 minute drive subscription about 16 minutes before it expires.
- Updates and deletes send `expected_etag` as If-Match; 412 maps to
  `ABORTED`. 429 and 503 map to `RESOURCE_EXHAUSTED` / `UNAVAILABLE` with
  `RetryInfo`.
- `Sync` with `include_security` re-reads security in full each pass. A
  resumed pass reads the site's change log from
  `SyncCursor.security_change_tokens` and emits again the items below a
  `SCOPE_ADD` or `SCOPE_DELETE` (`MicrosoftCrawler.rescopedOf`); a cursor
  older than the log's retention can miss one, and a full pass heals it.

## Review hot spots

- `MicrosoftCrawler`: scoped syncs treated as full sweeps, 410 on one
  drive's delta link, a full Sync without lists or security reconciling
  those rows as deleted (issues #10, #14, #23). Ledger write failures now
  fail the `Sync` (PR #20); check new sinks follow the same rule.
- `GraphPager`: silent fallback to Graph's default projection on 400 (#16).
- Output and OKF sinks shared across connections (#17), file names stored
  as payloads (#18).
- `grpc-microsoft-connector`: dropped content and mismatched delete ids
  (#12, #13).
- `SiteSecurityReader` / `SiteSecurityMapper`: Limited Access, sharing
  links, completeness when a read is refused (#3).
- `UploadHandler` resume tokens (HMAC-signed) and `MutateHandler` batching.
- `RecordingScrubber`: anything that could leak tenant data.
