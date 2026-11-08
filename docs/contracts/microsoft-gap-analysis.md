# Microsoft Graph gRPC contract: gap analysis and proposed proto changes

Scope: `grpc-microsoft-api/src/main/proto/ai/pipestream/microsoft/v1/` (the
`MicrosoftService` read proxy over Microsoft Graph v1.0 for SharePoint and
OneDrive), compared with what Graph v1.0 offers today for document-repository
crawling and indexing. Status: draft for review (Opus reviewer, then Kristian).

Sources: the Graph v1.0 OpenAPI document (`components/schemas/microsoft.graph.*`
and the `/sites`, `/drives`, `/shares`, `/users`, `/groups`, `/subscriptions`
paths), learn.microsoft.com for behaviour the OpenAPI does not carry (delta
semantics and `Prefer` headers, `getAllSites`, `sites/delta`, permission
roles, subscriptions, the "discover and detect changes at scale" guidance),
the current protos, and the Java implementation (`GraphFiles`, `GraphClient`,
`MicrosoftMapper`, `MicrosoftCrawler`, `MicrosoftGrpcService`,
`MicrosoftValidator`) to see what is actually called.

## 1. Summary

The current contract models four entities (Site, Drive, DriveItem, GraphUser)
with a handful of fields each, one folder-walk crawl, and no change tracking,
ACLs, lists, versions or schema. Concretely:

| Finding | Evidence | Consequence for an indexer |
|---|---|---|
| `Sync` never calls `/delta`. It is a recursive `/children` walk; `ChangeSource.CHANGE_SOURCE_DELTA` is declared but never emitted. | `MicrosoftCrawler.crawlFolder` -> `GraphFiles.childrenAll`; no `delta` anywhere in `grpc-microsoft-service`. | Every pass is a full crawl; deletes are never observed; `/children` paging is not guaranteed complete while writes happen (Graph says use delta for a mirror). |
| `SyncResponse.resume_cursor` is `Instant.now().toString()` and `SyncRequest` has no way to pass it back. | `MicrosoftGrpcService.sync`, `MicrosoftCrawler.crawl` (`setCursor(Instant.now().toString())`). | The "resume cursor" resumes nothing. |
| Inline content is capped at a hard-coded 25 MiB default and refused above it. | `DEFAULT_ATTACHMENT_MAX_BYTES` in crawler and service; the proto comments advertise "default 25 MiB". | Large documents cannot be fetched through the contract at all. |
| `ListSites` returns the first Graph page only. | Proto comment on `ListSitesRequest.limit`; `listSites` reads one `/sites?search=` page. | Tenants with more than one page of sites are silently truncated. |
| Every item costs two Graph calls (`/items/{id}` then `/items/{id}/listItem?$expand=fields`), always. | `MicrosoftCrawler.crawlFolder`, `getItem`, `downloadItem`. | Doubles throttling exposure for callers that do not need columns. |
| No permissions / sharing data. | No `permission` route in `GraphFiles`; no message. | No security trimming is possible. |
| No lists, list items, columns or content types beyond the flattened `list_columns` on a document. | No `/lists` route. | Non-library lists (tasks, custom lists, calendars) are invisible; column values have no schema. |
| No versions, no `sharepointIds`, no `remoteItem`/`shared`/`deleted`/`malware`/`package` facets, no `cTag`. | `MicrosoftMapper.toDriveItem`. | Cannot skip unchanged binaries (`cTag`), cannot de-duplicate shortcuts, cannot join to SharePoint search/audit (`sharepointIds`), may download malware. |
| The crawler calls `/drives/{id}` but there is no `GetDrive` rpc; there is no `GetSite`, no `/sites/getAllSites`, no `/users` discovery for OneDrives. | `GraphFiles.drive`; service rpc list. | App-only tenant-wide discovery is not expressible. |
| `Drive.site_id` is set only when the drive was listed through a site. | `MicrosoftMapper.toDrive(node, siteId)`. | Drive-to-site joins are lossy; Graph's `parentReference.siteId`/`sharePointIds.siteId` are not mapped. |
| `FileHashes.sha256` is a property Graph documents as "not supported, don't use". | OpenAPI `hashes.sha256Hash` description. | Kept for wire compatibility, comment corrected. |

The proposal (section 4) closes these with delta rpcs for drives, lists and
sites (deletes and resumable `@odata.deltaLink` cursors, 410 resync surfaced
as a typed event), a permission model (`sharePointIdentitySet`, links,
inheritance, per-item and per-site), lists / list items / columns / content
types, versions, the full set of driveItem facets, a streaming download rpc,
tenant-wide discovery (`getAllSites`, `/users`), and request-level include
flags so the proxy only pays for expansions the caller asked for. Everything is
additive; the descriptor diff against the current protos shows 0 changed or
removed fields, enum values, or rpc signatures (section 7).

## 2. Current contract vs Graph v1.0 (entities and fields)

Legend: present = modelled today; partial = modelled but lossy or wrong;
missing = not modelled. "Proposed" names the message/field added.

### 2.1 site (`microsoft.graph.site`)

| Graph property | Current | Proposed | Why |
|---|---|---|---|
| id, name, displayName, webUrl | present (`Site` 1-4) | kept | |
| description, createdDateTime, lastModifiedDateTime, eTag | missing | `Site.description/created_at/last_modified_at/etag` (5,6,7,10) | baseItem fields; freshness and change detection |
| createdBy, lastModifiedBy | missing | `Site.created_by/last_modified_by` (8,9) | rarely populated by Graph, cheap to carry |
| isPersonalSite | missing | `Site.is_personal_site` (11) | distinguish OneDrive sites in `getAllSites` output |
| root (facet) | missing | `Site.is_root` (12) | root vs subsite |
| siteCollection (hostname, dataLocationCode, root, archivalDetails) | missing | `SiteCollection` + `Site.site_collection` (13) | multi-geo and archived collections |
| sharepointIds | missing | `Site.sharepoint_ids` (14) | join key to SharePoint search / audit |
| parentReference | missing | `Site.parent_reference` (15) | subsite -> parent |
| deleted (delta) | missing | `Site.deleted` (16) | sites delta deletes (see open question 6) |
| analytics, columns, contentTypes, drives, lists, sites, permissions (navigations) | missing | rpcs: `GetItemAnalytics`, `ListColumns`, `ListContentTypes`, `ListDrives`, `ListLists`, `ListSubsites`, `ListPermissions` | navigations become rpcs, not embedded collections |
| onenote, items, externalColumns, operations, pages, termStore | missing | left out (section 5) | |

### 2.2 drive (`microsoft.graph.drive`)

| Graph property | Current | Proposed | Why |
|---|---|---|---|
| id, name, driveType, webUrl | present | kept (`drive_type` stays an open string, as Graph declares it) | |
| site_id | partial (only when listed via site) | kept; server also fills it from `parentReference.siteId` / `sharePointIds.siteId` | lossless join |
| description, createdDateTime, lastModifiedDateTime, createdBy, lastModifiedBy, eTag | missing | `Drive` 6-10, 15 | baseItem fields |
| owner | missing | `Drive.owner` (11) | whose OneDrive / which group |
| quota | missing | `Quota` + `Drive.quota` (12) | capacity reporting |
| sharePointIds | missing | `Drive.sharepoint_ids` (13) | join key |
| system (facet) | missing | `Drive.is_system` (14) | skip system libraries (Site Assets, Form Templates) |
| list (navigation) | missing | `Drive.list_id` (16) + `GetList(drive_id)` | library <-> list join |
| root, items, special, following, bundles (navigations) | partial (`ListChildren`) | `GetItem(path or id)`, `ListChildren(folder_id or path)`, `DeltaDriveItems`; special/following/bundles left out | |

### 2.3 driveItem (`microsoft.graph.driveItem`)

| Graph property | Current | Proposed | Why |
|---|---|---|---|
| id, name, webUrl, size, description, eTag, createdDateTime, lastModifiedDateTime, createdBy, lastModifiedBy | present | kept | |
| @microsoft.graph.downloadUrl | present (`download_url`) | kept; comment notes it is short-lived | |
| file.mimeType, file.hashes | present (flattened) | kept; `sha256` comment corrected | |
| folder.childCount | present | kept | |
| parentReference | partial (`drive_id`, `parent_id`) | `DriveItem.parent_reference` (24, full `ItemReference`) | path, siteId, sharepointIds, shareId |
| cTag | missing | `DriveItem.ctag` (22) | skip unchanged binaries (Graph's own guidance) |
| webDavUrl | missing | `web_dav_url` (23) | |
| fileSystemInfo | missing | `FileSystemInfo` + `file_system_info` (25) | user-visible timestamps |
| deleted (delta facet) | missing | `Deleted` + `deleted` (26) | deletes in delta / Sync |
| root (facet) | missing | `is_root` (27) | |
| package | missing | `Package` + `package` (28) | OneNote notebooks are folders on the wire |
| remoteItem | missing | `RemoteItem` + `remote_item` (29) | shortcuts / shared-with-me pointers; index the target, not the pointer |
| shared | missing | `Shared` + `shared` (30) | sharing scope (anonymous / organization / users) |
| specialFolder | missing | `SpecialFolder` (31) | |
| image, photo, video, audio, location | missing | `Image`, `Photo`, `Video`, `Audio`, `GeoCoordinates` (32-36) | media metadata for indexing |
| malware | missing | `Malware` + `malware` (37) | never download flagged items |
| publication | missing | `PublicationFacet` + `publication` (38) | checked-out / published state |
| sharepointIds | missing | `SharepointIds` + `sharepoint_ids` (39) | join key (list id, list item id, unique id, site/web/tenant ids) |
| pendingOperations.pendingContentUpdate | missing | `pending_content_update` (40) | binary about to change |
| retentionLabel (navigation) | missing | `ItemRetentionLabel` + `retention_label` (41), fetched on `include.retention_label` | compliance metadata |
| listItem.contentType, listItem.id | missing | `content_type` (42), `list_item_id` (43) | join to list item rpcs |
| listItem.fields | present (`list_columns`, always fetched) | kept; fetched only on `include.list_columns` | halves Graph calls when not needed |
| permissions (navigation) | missing | `Permission` + `permissions` (44), fetched on `include.permissions` | ACL for security trimming |
| @microsoft.graph.sharedChanged (delta annotation) | missing | `sharing_changed` (45) | re-read ACL only where it changed |
| analytics (navigation) | missing | `ItemAnalytics` + `analytics` (46), fetched on `include.analytics` | popularity signal |
| versions (navigation) | missing | `DriveItemVersion` + `ListItemVersions` rpc; version content via `DownloadContent.version_id` | |
| content | present (`optional bytes content = 21`, 25 MiB cap) | kept, cap removed (request `max_inline_bytes`, 0 = no limit); new streaming `DownloadContent` | no hard-coded caps |
| searchResult, bundle, thumbnails, workbook, subscriptions, activities | missing | left out (section 5); subscriptions handled at drive-root level | |

### 2.4 list, listItem, columnDefinition, contentType (all missing today)

| Graph resource | Proposed message | Graph route(s) | Notes |
|---|---|---|---|
| list (+ listInfo, system, sharepointIds, drive) | `SharePointList`, `ListInfo` | `/sites/{s}/lists`, `/sites/{s}/lists/{l}`, `/drives/{d}/list` | named `SharePointList` so the generated Java class does not collide with `java.util.List`; field names keep Graph's `list` |
| listItem (+ contentType, fields, sharepointIds, deleted, driveItem, permissions, analytics, versions) | `ListItem`, `ListItemVersion` | `/sites/{s}/lists/{l}/items[?$expand=fields]`, `/items/{i}`, `/items/delta`, `/items/{i}/versions` | `fields` reuse the existing typed `ListColumn`; `ListColumn.string_values` added so multi-value columns stop being lossy |
| columnDefinition (+ every column facet, defaultValue, validation, sourceContentType, columnTypes enum) | `ColumnDefinition`, `TextColumn`, `ChoiceColumn`, `LookupColumn`, `PersonOrGroupColumn`, `DateTimeColumn`, `NumberColumn`, `CurrencyColumn`, `CalculatedColumn`, `HyperlinkOrPictureColumn`, `TermColumn`, `BooleanColumn`, `GeolocationColumn`, `ThumbnailColumn`, `ContentApprovalStatusColumn`, `DefaultColumnValue`, `ColumnValidation`, `DisplayNameLocalization`, enum `ColumnType` | `/sites/{s}/columns`, `/sites/{s}/lists/{l}/columns`, `/sites/{s}/contentTypes/{ct}/columns` | facets are a `oneof kind`; `columnTypes` is a closed OData enum so it is a protobuf enum |
| contentType (+ columnLinks, inheritedFrom, columns) | `ContentType`, `ColumnLink` | `/sites/{s}/contentTypes`, `/sites/{s}/lists/{l}/contentTypes` | documentSet / documentTemplate / order left out |

### 2.5 permission (missing today)

| Graph property | Proposed | Notes |
|---|---|---|
| id, roles | `Permission.id`, `roles` (repeated string) | roles is an open string collection in Graph (`read`, `write`, `owner`, SharePoint `sp.*` and site `manage`/`fullcontrol`); kept verbatim |
| grantedToV2, grantedToIdentitiesV2 | `granted_to`, `granted_to_identities` (`SharePointIdentitySet`) | deprecated `grantedTo`/`grantedToIdentities` are not modelled; the server folds them into the V2 fields |
| sharePointIdentitySet (user, group, application, device, siteUser, siteGroup, sharePointGroup) | `SharePointIdentitySet`, `SharePointIdentity` (loginName), `SharePointGroupIdentity` | directory identities for trimming; site-scoped principals for SharePoint-only groups |
| link (type, scope, webUrl, webHtml, preventsDownload, application) | `SharingLink` | `scope = anonymous` is the "public" signal |
| invitation | `SharingInvitation` | |
| inheritedFrom | `inherited_from` (`ItemReference`) | not returned by ODB / SharePoint libraries |
| shareId, expirationDateTime, hasPassword | `share_id`, `expires_at`, `has_password` | DateTime.MinValue -> unset |
| (no back-reference in Graph) | `PermissionTarget target` | lets a permission travel alone (Sync change, `ListPermissions`) |

### 2.6 Other resources

| Graph resource | Current | Proposed |
|---|---|---|
| user (`/me`, `/users`) | `GraphUser` (4 fields) | + given_name, surname, job_title, department, office_location, user_type (Member/Guest), account_enabled, preferred_language; `GetUser`, `ListUsers` |
| itemAnalytics / itemActivityStat / itemActionStat / incompleteData | missing | `ItemAnalytics`, `ItemActivityStat`, `ItemActionStat`, `IncompleteData`; `GetItemAnalytics` |
| itemRetentionLabel / retentionLabelSettings / behaviorDuringRetentionPeriod | missing | `ItemRetentionLabel`, `RetentionLabelSettings`, enum `RetentionBehavior` |
| driveItemVersion / listItemVersion / publicationFacet | missing | `DriveItemVersion`, `ListItemVersion`, `PublicationFacet` |
| subscription | missing | `Subscription`; `ListSubscriptions`, `CreateSubscription`, `RenewSubscription`, `DeleteSubscription` (optional, open question 1) |
| sharedDriveItem (`/shares/{id}`) | missing | left out (section 5) |

## 3. Current rpcs vs Graph routes

| Rpc | Today | Problem | Proposal |
|---|---|---|---|
| `GetMe` | `GET /me` | fails on app-only | kept; `connection_id` added; `ProbeConnection` is the mode-agnostic probe |
| `ListSites` | `GET /sites?search=` first page | truncation | kept unary; `limit = 0` now means walk every page; tenant enumeration moves to `ListAllSites` |
| `ListDrives` | `/sites/{id}/drives` or `/me/drive` | no user/group scope; `/me` only | + `user_id` (`/users/{id}/drives`), `group_id` (`/groups/{id}/drives`); empty scope = `/me/drives` |
| `ListChildren` | `/drives/{id}/root:/path:/children` | path only; no expansions | + `folder_id`, `include` |
| `GetItem` | `/drives/{id}/items/{id}` + always `/listItem?$expand=fields` | second call always paid | + `path`, `include` (list columns opt-in) |
| `DownloadItem` | unary, 25 MiB cap | hard cap; bulk bytes on a unary | deprecated (`option deprecated = true`); `max_inline_bytes` (0 = no limit) replaces the cap; `DownloadContent` streams |
| `Sync` | recursive children walk, timestamp cursor, no deletes | not incremental | delta-driven; `since_cursor`, `max_inline_bytes`, `include`, `site_ids`, `include_lists`, `delta_preferences`, `include_system`; `resume_cursor` becomes an encoded `SyncCursor` |

## 4. Proposed rpc surface (each with its Graph route and auth mode)

A = application (client credentials) token, D = delegated token.

| Rpc | Graph route | A | D | Streams |
|---|---|---|---|---|
| `GetMe` | `GET /me` | no | yes | |
| `ProbeConnection` | `GET /sites?search=*` (D) / `GET /sites/getAllSites` (A) | yes | yes | |
| `GetUser` | `GET /users/{id}` | yes | yes | |
| `ListUsers` | `GET /users[?$filter]` | yes | yes (limited) | yes |
| `ListSites` | `GET /sites?search={q}` | yes | yes | |
| `ListAllSites` | `GET /sites/getAllSites` | yes | no | yes |
| `GetSite` | `GET /sites/{id}` or `/sites/{host}:/{path}` | yes | yes | |
| `ListSubsites` | `GET /sites/{id}/sites` | yes | yes | yes |
| `DeltaSites` | `GET /sites/delta[?token]` | yes | no | yes |
| `ListDrives` | `GET /sites/{id}/drives`, `/users/{id}/drives`, `/groups/{id}/drives`, `/me/drives` | yes (not `/me`) | yes | |
| `GetDrive` | `GET /drives/{id}` | yes | yes | |
| `ListChildren` | `GET /drives/{id}/root:/{path}:/children`, `/items/{id}/children` | yes | yes | yes |
| `GetItem` | `GET /drives/{id}/items/{id}`, `/root:/{path}` | yes | yes | |
| `DownloadItem` (deprecated) | `GET .../content` | yes | yes | |
| `DownloadContent` | `GET /drives/{id}/items/{id}/content`, `.../versions/{v}/content`, `?format=` | yes | yes | yes |
| `DeltaDriveItems` | `GET /drives/{id}/root/delta`, `/items/{id}/delta`, `?token=latest`, `Prefer:` headers | yes | yes (sharing prefs need Sites.FullControl.All) | yes |
| `ListItemVersions` | `GET /drives/{id}/items/{id}/versions` | yes | yes | yes |
| `GetItemAnalytics` | `GET .../items/{id}/analytics`, `/sites/{id}/lists/{l}/items/{i}/analytics`, `/sites/{id}/analytics` | yes | yes | |
| `ListLists` | `GET /sites/{id}/lists` | yes | yes | yes |
| `GetList` | `GET /sites/{id}/lists/{l}`, `/drives/{d}/list` | yes | yes | |
| `ListListItems` | `GET /sites/{id}/lists/{l}/items?$expand=fields[&$filter]` | yes | yes | yes |
| `GetListItem` | `GET /sites/{id}/lists/{l}/items/{i}?$expand=fields` | yes | yes | |
| `DeltaListItems` | `GET /sites/{id}/lists/{l}/items/delta[?token]` | yes | yes | yes |
| `ListListItemVersions` | `GET /sites/{id}/lists/{l}/items/{i}/versions` | yes | yes | yes |
| `ListColumns` | `GET /sites/{id}/columns`, `/lists/{l}/columns`, `/contentTypes/{ct}/columns` | yes | yes | yes |
| `ListContentTypes` | `GET /sites/{id}/contentTypes`, `/lists/{l}/contentTypes` | yes | yes | yes |
| `ListPermissions` | `GET .../items/{id}/permissions`, `/lists/{l}/items/{i}/permissions`, `/lists/{l}/permissions`, `/sites/{id}/permissions` | yes (site perms need Sites.FullControl.All) | yes (not site perms) | yes |
| `Sync` | composition of the delta routes above | yes | yes | yes |
| `ListSubscriptions` / `CreateSubscription` / `RenewSubscription` / `DeleteSubscription` | `GET/POST /subscriptions`, `PATCH/DELETE /subscriptions/{id}` | yes | yes | |

## 5. Deliberately left out (and why)

| Graph capability | Reason |
|---|---|
| Thumbnails (`/thumbnails`), preview, `searchResult`, `bundle`, `/special`, `/following`, `/bundles`, `/recent`, `/sharedWithMe` | UI / per-user navigation, not repository indexing. `sharedWithMe` and `following` are delegated-only and the targets are reached by crawling the owning drives. |
| Drive / list search (`/search(q=)`) | A query surface, not a crawl surface; the ConnectorService being designed elsewhere decides what search looks like. |
| `/shares/{shareId}` resolution (`sharedDriveItem`) | Sharing links appear on `Permission.link`; resolving a link to an item is a client concern. Cheap to add later as `GetSharedItem`. |
| Workbook, OneNote, termStore, sitePages, `site.items`, `externalColumns`, long-running `operations` | Out of the document-repository scope. |
| `documentSetVersions`, `contentType.documentSet` / `documentTemplate` / `order`, `columnPositions`, `baseTypes` | Document-set and authoring metadata; not needed to index items. |
| Sensitivity labels (`extractSensitivityLabels`, `assignSensitivityLabel`) | POST actions that run label extraction server-side per file (metered); the label is also exposed through list columns (`_ComplianceTag` / sensitivity columns). Candidate for a later additive rpc. |
| Group membership expansion (`/groups/{id}/transitiveMembers`, `/users/{id}/transitiveMemberOf`) | Directory, not SharePoint. Needed by a trimming layer, but it belongs to the connector/identity design, not this vendor proxy. Note: SharePoint site-group membership is not available through Graph at all (SharePoint REST only). See open question 4. |
| Write operations (upload, create link, invite, check-in/out, copy, move, delete, restore, `PATCH listItem/fields`) | Read proxy. `GraphFiles` has upload helpers for the OKF publisher, but they are not on the gRPC surface and this proposal does not add them. |
| Deprecated `permission.grantedTo` / `grantedToIdentities` | Superseded by V2; the server maps old responses into the V2 fields. |
| `@odata.nextLink`, `@odata.count`, `@odata.deltaLink` as message fields | Transport. Lists stream; the delta link is the typed terminal event of delta streams and the `MicrosoftSnapshot.cursor` of Sync. |
| GCA `Contracts/` protos | Frozen Microsoft copy; untouched. |

## 6. Design decisions

1. **File split.** `common.proto` keeps identities and shared facets; entities
   move to `site.proto`, `drive.proto`, `list.proto`, `permission.proto`,
   `subscription.proto`; the crawl envelopes (`MicrosoftEntity`,
   `ChangeSource`, `ChangeOperation`, `MicrosoftSnapshot`, `MicrosoftChange`)
   move to `events.proto` (mirrors `confluence/v1/events.proto`). Moving was
   necessary: `MicrosoftEntity` must wrap `ListItem`, `Permission`, ... which
   import `common.proto`, so leaving the envelope in `common.proto` would be
   an import cycle. Fully qualified names, field numbers and the generated
   Java classes (`java_multiple_files = true`, no `java_outer_classname`, no
   Java code references an outer class) are unchanged. CI runs `buf lint`
   only; `buf breaking` with the `FILE` category would flag the moves, the
   `WIRE_JSON`/`WIRE` categories would not. Fallback if the move is
   unwelcome: keep every message in `common.proto` (same wire, one large
   file).
2. **Open strings stay strings.** Graph declares `driveType`, link
   `type`/`scope`, `roles`, `listInfo.template`, `quota.state`,
   `publication.level`, `deleted.state`, `shared.scope` as plain strings and
   adds values without notice. Only closed OData enums (`columnTypes`,
   `behaviorDuringRetentionPeriod`) became protobuf enums (with Graph's
   `unknownFutureValue` carried as a value).
3. **Delta modelling.** `DeltaDriveItems`, `DeltaListItems`, `DeltaSites`
   are 1:1 with Graph's `delta` functions: empty cursor = full enumeration,
   `latest` = cursor only, otherwise changes since the cursor; deleted items
   carry the `deleted` facet; the stream ends with the new `@odata.deltaLink`.
   HTTP 410 Gone is a typed first event (`DeltaResync` with Graph's
   `resyncChangesApplyDifferences` / `resyncChangesUploadDifferences`)
   followed by the fresh enumeration Graph redirects to, so the client can
   reconcile deletes against local state. `Prefer` headers are typed booleans
   (`DeltaPreferences`); `@microsoft.graph.sharedChanged` is
   `DriveItem.sharing_changed`.
4. **Sync becomes delta-driven but keeps its shape.** `SyncResponse` is
   unchanged; `MicrosoftSnapshot.cursor` now carries the per-drive / per-list
   delta link, and the terminal `resume_cursor` is the base64 of a typed
   `SyncCursor` (per-drive and per-list delta links, plus the sites delta
   link) so clients stay opaque and the server never needs JSON.
   `SyncRequest.since_cursor` hands it back. `DELETE` changes finally appear
   (`ChangeSource.CHANGE_SOURCE_DELTA` becomes real).
5. **No payload caps.** `DriveItem.content` stays for inline mode, governed
   by `max_inline_bytes` per request (0 = no limit). `DownloadContent`
   streams metadata first, then chunks, with optional range and chunk size.
   `DownloadItem` is deprecated, not removed.
6. **Expansions are opt-in.** `DriveItemInclude` / `ListItemInclude` make
   list columns, permissions, retention label, analytics, publication and
   drive-item expansion explicit per request (one Graph call or `$expand`
   each). Default-off changes the current always-on list-column fetch; see
   open question 2.
7. **Permissions travel with the item and alone.** `DriveItem.permissions`,
   `ListItem.permissions`, `SharePointList.permissions` for callers that want
   the ACL on the document; `Permission.target` + `MicrosoftEntity.permission`
   for callers that want ACL changes as their own events.
8. **Server-filled back-references** (`ListItem.site_id/list_id`,
   `DriveItemVersion.drive_id/item_id`, `ColumnDefinition.site_id/list_id/
   content_type_id`, `Permission.target`) are the only fields with no Graph
   counterpart; each is commented "set by the server".
9. **`SharePointList`** is the one naming deviation from Graph, for Java
   ergonomics (`java.util.List`). Field names keep Graph's `list`.
10. **`connection_id`** is added to every request (new and existing) to
    match the Confluence surface and the catalog `ConnectionService`.

## 7. Wire compatibility and validation evidence

* `protoc` (libprotoc 35.1, via `grpcio-tools`) compiles all eight files with
  `--include_imports` (`buf` was not available in the sandbox; the Gradle
  protobuf plugin uses protoc 4.29.3, which accepts the same proto3 syntax).
  Java generation was NOT run (the pip protoc ships without
  `protoc-gen-java`); generated-name clashes were reviewed by inspection
  (no `foo` / `foo_list` / `foo_count` pairs, no camelCase collisions, the
  `List` -> `SharePointList` rename). `./gradlew :grpc-microsoft-api:build`
  is the first thing to run on a branch.
* Descriptor diff of the pre-change descriptor set against the new one, over
  every message, field (number, name, type, label, proto3 optional, oneof
  membership), enum value and rpc (input/output type, streaming): 0
  incompatibilities; 40 -> 206 messages, 2 -> 6 enums, 7 -> 32 rpcs.
* A textual approximation of buf `COMMENTS` + `STANDARD` (leading comment on
  every service, rpc, message, oneof, field, enum, enum value; `*_UNSPECIFIED`
  zero values with enum-name prefixes; `<Rpc>Request`/`<Rpc>Response` naming;
  lower_snake_case fields): 0 problems. Real `buf lint` still needs to run in
  CI.
* Existing wire: `DriveItem.content = 21` keeps its number and `optional`
  label; `SyncResponse` arms unchanged; `MicrosoftEntity` arms 10-13 unchanged.

## 8. Implementation impact (informational; no Java touched here)

* `MicrosoftCrawler` should move to `/drives/{id}/root/delta` for both the
  first full pass and incremental passes (Graph's own recommendation for a
  mirror), using `$select=*,sharepointIds` and `$expand=listItem($expand=fields)`
  when `include.list_columns` is set (verify the `$expand` on delta against a
  live tenant; otherwise fall back to the per-item `/listItem` call).
* Deletes: emit `CHANGE_OPERATION_DELETE` with an identity-only `DriveItem`
  carrying `deleted`; the sync-table sink's `Reconcile` becomes unnecessary
  for delta passes.
* `MicrosoftValidator` needs cases for the new entity arms (`SharePointList`,
  `ListItem`, `Permission`, versions, `ColumnDefinition`, `ContentType`), and
  `MicrosoftEntity.entity_id` for a permission should be the permission id.
* Remove `DEFAULT_ATTACHMENT_MAX_BYTES`; honour `max_inline_bytes`; implement
  `DownloadContent` by streaming the `/content` (or `download_url`) response
  body in `chunk_bytes` pieces, honouring `Range`.
* `GraphClient` already retries 429/503 with `Retry-After`; delta streams
  should pause all work during back-off (Graph throttles per app per tenant).
* 410 handling: on `GET delta` returning 410, read the `Location` header and
  the error `code`, emit `DeltaResync`, continue from `Location`.
* Permissions with `hierarchicalsharing` + `deltashowsharingchanges` need
  `Sites.FullControl.All` (application); document in the connector setup.

## 9. Open questions for Kristian

1. **Subscriptions on this surface?** Managing Graph webhook subscriptions is
   a Graph call and the only way to stop timer-polling delta, but receiving
   the notification is an HTTPS endpoint outside gRPC, which the
   ConnectorService design may own. Keep the four `*Subscription` rpcs here,
   or drop `subscription.proto` until that design lands?
2. **Default for list columns.** Today `GetItem` / `Sync` always fetch
   `listItem.fields` (an extra call per item). The proposal makes it opt-in
   (`include.list_columns`). Accept the behaviour change, or default it on
   in the server for `Sync` to keep the OKF / connector output identical?
3. **File split vs. one `common.proto`.** Section 6.1: the move is
   wire-identical and Java-identical but `buf breaking` (FILE) would flag it.
   OK to switch `buf.yaml` `breaking.use` to `WIRE_JSON`, or keep everything
   in `common.proto`?
4. **Group expansion for trimming.** Should the Microsoft proxy expose
   `/groups/{id}/transitiveMembers` and `/users/{id}/transitiveMemberOf`
   (directory routes, app-only `GroupMember.Read.All`), or does ACL expansion
   live with the connector / identity layer? SharePoint site-group membership
   is not available through Graph in any case.
5. **`Sync` scope for lists.** `include_lists` sweeps non-library lists via
   list-items delta. Should library lists also be emitted as `ListItem`
   changes (duplicating `DriveItem`), or only as drive items (current
   proposal)?
6. **Sites delta deletes.** Graph's `sites/delta` documentation does not
   show an explicit delete marker; `Site.deleted` is modelled on the
   driveItem/listItem pattern. Needs a live-tenant check before `DeltaSites`
   is relied on for site removal; `ListAllSites` reconciliation is the
   fallback.
7. **`ListSites` semantics.** The proposal reinterprets `limit = 0` as "walk
   every page" (the comment previously said "first page only"). Fine as a
   comment fix, or should the unary rpc keep a server-side page bound?
8. **`DriveItem.content` long term.** Keep inline content for `Sync`
   (connector and OKF sinks read it today), or move sinks to
   `DownloadContent` and deprecate field 21 in a later revision?
9. **Analytics / retention label** are in the proposal as opt-in expansions
   (one extra call each). Worth the surface, or defer?
10. **`ProbeConnection` inline secrets.** Mirrors Confluence's inline
    `token`; confirm the catalog `ConnectionService` is the intended owner of
    unsaved-credential probes and this is acceptable on the Microsoft side.

## 10. Files

Proto (edited in place, all under
`/home/claude/grpc-confluence/grpc-microsoft-api/src/main/proto/ai/pipestream/microsoft/v1/`):

* `common.proto` (rewritten: shared identities, refs, facets; entities moved out)
* `site.proto` (new)
* `drive.proto` (new: Drive, DriveItem and facets, DriveItemVersion, retention)
* `list.proto` (new: SharePointList, ListItem, ListItemVersion, columns, content types)
* `permission.proto` (new)
* `subscription.proto` (new)
* `events.proto` (new: MicrosoftEntity, ChangeSource, ChangeOperation, MicrosoftSnapshot, MicrosoftChange, SyncCursor)
* `microsoft_service.proto` (rewritten: 32 rpcs, request options, delta events)

Scratch tooling used for validation:
`scratchpad/tools/compat.py`
(descriptor diff), `.../tools/commentlint.py` (comment / naming lint),
`.../tools/out/baseline.pb` and `.../tools/out/new.pb` (descriptor sets).

## Review (Opus)

Reviewed the draft (commit `9098195`, compared with `HEAD~1`) against the
Graph v1.0 OpenAPI, the existing Java and the repo conventions. Edits are
applied in the working tree, not committed.

### Build status

* `bash ./gradlew -q :grpc-microsoft-api:compileJava`: passes.
* `compileJava` + `compileTestJava` for `grpc-microsoft-service`,
  `grpc-microsoft-connector`, `grpc-okf`, `grpc-output-spi`,
  `grpc-output-filesystem`, `grpc-output-s3`, `grpc-mcp`: pass.
* `grpc-connect` depends on `grpc-confluence-service`. That module currently
  fails in the shared tree because the parallel Confluence review is
  mid-edit (`ConfluenceValidator` switch). Compiled in a scratch copy with
  `HEAD~1` Confluence sources plus these Microsoft changes:
  `:grpc-connect:compileJava :grpc-connect:compileTestJava` pass.
* Descriptor diff `HEAD~1` -> working tree (protoc via grpcio-tools, the
  draft's `compat.py`): 0 incompatibilities. 7 -> 23 rpcs, 40 -> 162
  messages, 2 -> 6 enums. The approximate buf STANDARD+COMMENTS lint
  (`buflint.py` pointed at the microsoft package, plus `commentlint.py`)
  reports 0 problems. Real `buf lint` still has to run in CI.

### Build and Java fixes

1. **Outer-class collisions.** protoc had quietly generated
   `DriveOuterClass`, `SiteOuterClass`, `PermissionOuterClass` and
   `SubscriptionOuterClass`. Worse, it generated an outer class
   `ai.pipestream.microsoft.v1.List` for `list.proto`, which shadows
   `java.util.List` for any code that imports the package. The new files now
   set `java_outer_classname`: `DriveProto`, `SiteProto`, `ListProto`,
   `PermissionProto`, `EventsProto`. `common.proto` (`Common`) and
   `microsoft_service.proto` (`MicrosoftServiceOuterClass`) are unchanged.
   No Java references an outer class, and every message class keeps its
   name and package (`java_multiple_files`), so moving the envelopes to
   `events.proto` does not affect Java.
2. **New `MicrosoftEntity` arms broke exhaustive switches.** The draft said
   it touched no Java, but `grpc-okf` `MicrosoftCatalog.from` stopped
   compiling and `MicrosoftValidator.visitEntity` was not exhaustive. Added
   `LIST`, `LIST_ITEM` and `PERMISSION` cases (catalog fallback entry;
   validator visits the nested message). `ObjectKeys` already has a
   `default`.
3. Renamed the facet message `Package` to `PackageFacet` (it shadows
   `java.lang.Package` under on-demand imports). The field name is still
   `package`, matching Graph. The message is new, so this is wire-neutral.

### Correctness fixes

4. **Folder-scoped delta.** `/items/{id}/delta` on a non-root folder works
   on OneDrive personal only. OneDrive for Business and SharePoint support
   delta on the drive root only. Fixed `DeltaDriveItems`,
   `DeltaDriveItemsRequest.item_id` (INVALID_ARGUMENT on business drives)
   and `SyncRequest.folder_path` (a delta pass covers the whole drive, and
   the path only filters the output). The draft said Graph "scopes delta to
   a folder by item id", which is wrong for SharePoint.
5. **Delta links are URLs supplied by the client.** If the server followed
   them with its bearer token, that would be an SSRF and token-exfiltration
   path. Every `delta_link` field and `SyncCursor` now states that the
   server follows only links on its configured Graph host and the matching
   delta route, and answers INVALID_ARGUMENT for anything else.
6. **Sync deletes everything after an incremental pass.**
   `SyncTableMicrosoftChangeSink.completeRun` reconciles by run id, so it
   deletes every asset the run did not touch. That was correct for full
   crawls. On a delta-driven incremental pass it would delete every
   unchanged asset. Added `MicrosoftSnapshot.source` (`ChangeSource`, field
   9): `CRAWL` means a full enumeration (first pass, children walk, or the
   fresh enumeration after a 410), so reconciling is allowed. `DELTA` means
   an incremental pass, so reconciling is forbidden. `UNSPECIFIED` means a
   first-revision server, which only did full sweeps. The implementation
   must gate `completeRun`'s reconcile on this. `SyncRequest.since_cursor`
   documents the 410 path.
7. **Silent behaviour change on list columns.** Today `GetItem`,
   `DownloadItem` and `Sync` always fetch `listItem.fields`. The draft
   switched that off by default, so existing clients (connector, OKF, MCP)
   would quietly lose columns. Now, when the `include` message is unset,
   each rpc keeps its first-revision behaviour: list columns on for
   `GetItem`, `DownloadItem`, `DownloadContent` and `Sync`, nothing extra
   for `ListChildren` and `DeltaDriveItems`. When `include` is set, exactly
   the flags set to true are fetched. Proto3 message presence makes this
   unambiguous, and it resolves open question 2 without a breaking change.
8. **Permission entity ids.** An inherited grant repeats the same
   permission id on every item below it, so a bare permission id is not a
   unique `entity_id`. The arm comment now requires target key plus
   permission id. Section 8's "entity_id should be the permission id" is
   wrong.
9. Removed `SharePointList.permissions`: no request could populate it.
   `ListPermissions` with a `list` target covers it.
10. `Site.deleted`: Graph's `site` schema has no `deleted` property. The
    comment now says the field is unverified and that callers should
    reconcile with `ListAllSites` until a live-tenant check confirms it
    (open question 6 stands).
11. `pending_content_update` comment notes that `queuedDateTime` is dropped.

Checked and found correct: permission V2 shape (`grantedToV2`,
`grantedToIdentitiesV2`, `sharePointIdentitySet` with
`group`/`siteUser`/`siteGroup`/`sharePointGroup`, `sharingLink`,
`sharingInvitation`, `inheritedFrom`, `shareId`, `expirationDateTime`,
`hasPassword`); the `columnTypes` and `behaviorDuringRetentionPeriod` enums
match the closed OData enums exactly, `unknownFutureValue` included; the
`Prefer` header names and their dependency (`deltashowsharingchanges` needs
`deltashowremovedasdeleted` + `deltatraversepermissiongaps`); the 410
`resyncChangesApplyDifferences`/`UploadDifferences` codes; `cTag` vs `eTag`;
the Graph routes behind every kept rpc exist in the v1.0 OpenAPI (including
`lists/{id}/items/delta`, list and list-item `permissions`, and
`sites/{id}/permissions`, which returns application grants such as
`Sites.Selected`, not user ACLs). No `Struct`/JSON, no hard caps
(`max_inline_bytes` 0 = no limit, streaming `DownloadContent`), and
`metadata = 99` is preserved on every message that had it.

### Scope cuts (32 -> 23 rpcs, 206 -> 162 messages)

None of the cut items existed in `HEAD~1`, so removing them is wire-neutral.
All can come back later as additions. The field numbers were never released
and are not reserved.

| Cut | Why |
|---|---|
| `ListSubscriptions`, `CreateSubscription`, `RenewSubscription`, `DeleteSubscription`, `subscription.proto` | These are writes to Graph, and they only make sense next to an HTTPS webhook receiver, which this gRPC read proxy cannot host. The owner of the receiver (the connector / ConnectorService) should own the subscription lifecycle. Delta polling already gives a complete feed. **For Kristian:** confirm the connector owns webhooks. |
| `GetUser`, `ListUsers` | Directory, not SharePoint. App-only OneDrive discovery is already covered: `ListAllSites` returns personal sites (`is_personal_site`), then `ListDrives(site_id)`. `ListDrives.user_id` stays for a targeted lookup. ACL principal expansion belongs to the identity layer (open question 4). The enriched `GraphUser` fields stay for `GetMe`. |
| `GetItemAnalytics`, `DriveItemInclude.analytics`, `ListItemInclude.analytics`, `DriveItem.analytics`, `ListItem.analytics`, `ItemAnalytics`/`ItemActivityStat`/`ItemActionStat`/`IncompleteData` | A speculative ranking signal that costs one Graph call per item and is heavily throttled. Graph's analytics data is also known to be sparse. Nothing downstream consumes it. |
| `ListContentTypes`, `ContentType`, `ColumnLink` | Items already carry `ContentTypeInfo` (id + name). `ListColumns(content_type_id)` gives a content type's schema. Content-type authoring metadata is not needed to index items. |
| `ListListItemVersions`, `ListItemVersion` | Versions of generic-list items are rarely indexed. Drive-item versions stay (`ListItemVersions` + `DownloadContent.version_id`) because document versions have downloadable content. |
| `MicrosoftEntity` arms `drive_item_version`, `list_item_version`, `column_definition`, `content_type` | `Sync` has no request option that would ever emit them, so they were dead arms. The `list`, `list_item` and `permission` arms stay. |

Kept on purpose: the drive-item media facets (`image`, `photo`, `video`,
`audio`, `location`). They are properties on the entity, with no extra rpc
or Graph call, so vendor fidelity wins. Retention label stays as an opt-in
expansion because compliance metadata is a real enterprise-indexer need.
`ListAllSites`, `DeltaSites` and `ListSubsites` all stay: `getAllSites`
returns site collections only, subsites own libraries, and `sites/delta` is
the only cheap way to detect new sites.

### Remaining questions for Kristian

1. Subscriptions: drop here, owned by the connector (applied)? Or bring
   them back if the proxy is meant to manage webhooks?
2. `ProbeConnection` takes an inline `client_secret` over gRPC (it mirrors
   Confluence's inline token). Is it acceptable, or should unsaved-credential
   probes live only in the catalog `ConnectionService`? Kept as drafted.
3. `SyncCursor` travels as base64 in `resume_cursor` and grows with the
   number of drives and lists in scope (about 1 KB of delta link each). Is a
   multi-MB cursor acceptable for tenant-wide syncs, or should the server
   persist delta links (checkpoint store / sync table) and hand back a short
   key?
4. Original open questions 3 (file split vs `buf breaking` FILE; switching
   to `WIRE_JSON` is recommended), 4 (group expansion; recommend the identity
   layer), 5, 6, 7, 8 and 10 still stand. Question 2 is resolved by fix 7
   above. Question 9 is resolved for analytics (cut); retention label is kept.
5. A server answer to unbounded inline content (`max_inline_bytes = 0`) is
   still subject to the gRPC channel's max message size. Should Sync callers
   be steered to `DownloadContent` (open question 8)?

### Final rpc list (23)

Identity: `GetMe`, `ProbeConnection`.
Sites: `ListSites`, `ListAllSites`, `GetSite`, `ListSubsites`, `DeltaSites`.
Drives: `ListDrives`, `GetDrive`, `ListChildren`, `GetItem`, `DownloadItem`
(deprecated), `DownloadContent`, `DeltaDriveItems`, `ListItemVersions`.
Lists: `ListLists`, `GetList`, `ListListItems`, `GetListItem`,
`DeltaListItems`, `ListColumns`.
ACL: `ListPermissions`.
Crawl: `Sync`.
