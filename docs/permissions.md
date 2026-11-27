# Permissions the service identity needs

The least privilege each rpc needs in Microsoft 365, and how each row is
known. Update this file in the same pull request that adds an rpc or changes
what an identity must hold; `PermissionsDocTest` fails the build when an rpc
is not named here.

- **Proven**: passed, or was refused, in a live run against the test tenant.
- **Assumed**: from Microsoft's documentation or by inference; never tried.

Last checked against the test tenant: 2026-11-27.

## Identities

Two app registrations, both single tenant, application permissions with
admin consent. No signed-in user.

| Identity | Holds | Credential |
|---|---|---|
| Connector app (`MicrosoftService`, `MicrosoftChangeService`) | Graph `Sites.Selected`; SharePoint Online `Sites.Selected`; one grant per site (below). For Entra reads also Graph `User.Read.All` and `GroupMember.Read.All` | Certificate. A client secret is enough only while no rpc that calls SharePoint REST is used |
| Admin app (`MicrosoftAdminService`) | Graph `Sites.FullControl.All` | Certificate or client secret |

- Graph roles do not carry over to SharePoint REST: the SharePoint Online
  `Sites.Selected` role is granted separately (proven).
- SharePoint REST refuses an app-only token obtained with a client secret, so
  those rpcs answer `FAILED_PRECONDITION` without the certificate (proven).
- The site grant is given to the connector app by the admin service
  (`EnsureSiteAppGrant`), per site, with one role: `read`, `write`, `manage`
  or `fullcontrol`.
- Not needed: Graph `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.Read.All`.
  They were removed from the test app on 2026-11-22 and every site-scoped rpc
  still passes (run 45031978792).
- Consenting to application permissions needs a Global Administrator or a
  Privileged Role Administrator (assumed).

## Connector app: by rpc

Graph `Sites.Selected` plus the site grant in the second column. "REST" means
the rpc calls SharePoint REST and also needs SharePoint Online
`Sites.Selected` and the certificate.

| rpcs | Least site grant | How it is known |
|---|---|---|
| `ProbeConnection`, `GetSite`, `ListSubsites`, `ListDrives`, `GetDrive`, `ListChildren`, `GetItem`, `DownloadItem`, `DownloadContent`, `ListItemVersions`, `DeltaDriveItems`, `ListLists`, `GetList`, `ListListItems`, `GetListItem`, `DeltaListItems`, `ListColumns`, `Sync` without `include_security`, `ListPermissions` on a drive item, folder, list or list item | `read` | Assumed for `read` and `write`. Proven with `manage` (run 45031066801) and `fullcontrol` |
| `UploadItem`, `CreateFolder`, `UpdateDriveItem`, `DeleteDriveItem`, `RestoreDriveItemVersion`, `CopyDriveItem` | `write` | File writes passed with `write` (2026-11-21). The whole row is proven with `manage` |
| `CreateListItem`, `UpdateListItemFields`, `DeleteListItem`, `Mutate` | `write` | Assumed for `write`. Proven with `manage` |
| `EnsureList`, `EnsureColumns` | `manage` | Proven: 403 with `write`, pass with `manage` (run 45016824098) |
| Hyperlink and term columns in `EnsureColumns`, hyperlink values in any list item write (REST) | `manage` | Assumed for `manage`. Proven with `fullcontrol` (runs 45053446390, 45100103838). Graph refuses these, so they go through REST |
| `ListSiteGroups`, `GetSiteRoleInheritance` (REST) | `manage` | Proven with `manage` (2026-11-22) |
| `ListSiteGroupMembers`, `ListSiteRoleAssignments`, `Sync` with `include_security` (REST) | `fullcontrol` | Proven: refused with `manage`, pass with `fullcontrol` (run 45038404991) |
| `ListPermissions` on a site | `fullcontrol` | Proven: refused with `manage` (run 45031066801), pass with `fullcontrol` (run 45065737777) |
| `ListSiteChanges`, `WatchChanges` with `site_changes` (REST) | `fullcontrol` | Proven with `fullcontrol`. `manage` not tried |
| `SetSiteRoleInheritance`, `CreateSiteGroup`, `DeleteSiteGroup`, `AddSiteGroupMember`, `RemoveSiteGroupMember`, `ListSiteRoleDefinitions`, `AddSiteRoleAssignment`, `RemoveSiteRoleAssignment` (REST) | `fullcontrol` | Proven with `fullcontrol`. `manage` not tried |
| `CreateSubscription`, `GetSubscription`, `ListSubscriptions`, `UpdateSubscription`, `ReauthorizeSubscription`, `DeleteSubscription` on a drive | `read` | Assumed below `fullcontrol`. Proven with `fullcontrol`, security notifications included. Microsoft documents `Files.ReadWrite.All`; the site grant was enough. A subscription on a list was not tried |
| `WatchChanges` | What its `Sync` passes need (rows above) | Proven with `fullcontrol` |
| `WatchNotifications`, `ListSubscriptionRenewals` | Nothing in Microsoft 365: they read the server's own state | By construction |

Not tied to a site grant:

| rpcs | Needs | How it is known |
|---|---|---|
| `ListUsers`, `DeltaUsers` | Graph `User.Read.All` | Proven with both Entra permissions held (run 45031066801). Each alone not tried |
| `ListGroups`, `DeltaGroups`, `ListGroupTransitiveMembers`, `ListGroupOwners` | Graph `GroupMember.Read.All` | Same |
| `ListUserTransitiveMemberOf`, `WatchChanges` with `directory`, subscriptions on a user, a group or a group's members | Graph `User.Read.All` and `GroupMember.Read.All` | Proven |
| `GetGroup`, `CreateGroup`, `UpdateGroup`, `DeleteGroup`, `AddGroupMember`, `RemoveGroupMember` (on a branch, not merged) | Graph `Group.ReadWrite.All` | Not proven. Granted to the test app on 2026-11-27 |
| `ListSites`, `ListAllSites`, `DeltaSites` | Graph `Sites.Read.All` (tenant-wide) | Proven refused with `Sites.Selected`. Outside the minimal set: name the sites in the configuration instead |
| `GetTermStore`, `ListTermGroups`, `ListTermSets`, `ListTerms` | A signed-in user (delegated) | Not proven. Graph does not serve the term store to an application token (run 45095625449) |
| `GetMe` | A signed-in user (delegated) | Proven: `FAILED_PRECONDITION` with an application token |

So the smallest working sets are:

| Use | Site grant | Entra permissions | Certificate |
|---|---|---|---|
| Index content | `read` (assumed; `manage` proven) | none | no |
| Index content with its access control | `fullcontrol` | `User.Read.All`, `GroupMember.Read.All` | yes |
| Write content and metadata | `manage`; `fullcontrol` proven for hyperlink and term columns | none | for hyperlink and term columns |
| Change SharePoint groups and role assignments | `fullcontrol` | none | yes |

## Admin app

| rpcs | Needs | How it is known |
|---|---|---|
| `ListSiteAppGrants`, `EnsureSiteAppGrant`, `RevokeSiteAppGrant` | Graph `Sites.FullControl.All` | Proven: a grant was created, changed and revoked (runs 45037974601, 45043224505) |

The admin app has no Entra permission and cannot list users or groups
(proven). It runs as its own process and only when grants change.

## Open

- `read` and `write` grants for the read rpcs, and `manage` for the rows
  marked "not tried": each needs one live coverage run with the test app's
  grant lowered, then restored.
- Entra group writes need a live run before their row can say proven.
  `Group.ReadWrite.All` is tenant-wide: it lets the app change every group.
- The test admin app also holds SharePoint Online `Sites.FullControl.All`,
  used once to read role assignments as the admin. The admin service does not
  call SharePoint REST, so it should be removable (assumed).
- The test connector app's "configured permissions" list in the Entra portal
  does not match what is granted: it still names `Sites.Read.All`,
  `Sites.ReadWrite.All` and `Files.Read.All`, which are not granted. Do not
  press "Grant admin consent" there until the list is corrected; it would
  grant those three.
