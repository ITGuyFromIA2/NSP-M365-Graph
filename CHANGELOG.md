# Changelog

## 0.1.3 (2026-11-03)

- `Find-NSPGraphSite`: read-only site search, link resolution (shares API) and the site behind a
  library, returning site, drive and item IDs.
- `Get-NSPGraphDriveFolder`: read-only listing of a site's document libraries or a folder's
  subfolders.
- `New-NSPGraphDriveFolder`: creates one folder, failing rather than renaming on a name conflict.
- `Grant-NSPGraphSitePermission`: plan-first Sites.Selected site permission (read or write) for one
  app. It upgrades a weaker grant, never touches other apps' permissions, and leaves an equal or
  stronger grant alone.

## 0.1.2 (2026-11-02)

- Add optional app-only ContextScope forwarding and returned-scope validation. Omitting it
  preserves existing SDK scope selection. Add isolated offline regression coverage.
- `Find-NSPGraphMailbox`: read-only search, or exact address lookup, for mail-enabled member
  accounts, with a LikelyShared/LicensedUser hint from sign-in and license state.
- `Set-NSPGraphMailSendScope`: plan-first Exchange RBAC for Applications setup so an app sends as one
  mailbox only. Refuses while an unscoped Entra Mail.Send grant exists; verifies with
  Test-ServicePrincipalAuthorization.
- `Register-NSPGraphAppOnlyRegistration` accepts an empty `-ApplicationPermissions` list, for apps
  whose access is granted outside Graph.
- `Invoke-NSPGraphRequest` / `Invoke-NSPGraphCollection`: optional `-Headers` (for example
  ConsistencyLevel for `$search`).

## 0.1.1 (2026-10-28)

- `Send-NSPGraphMail`: one sendMail request as a mailbox through the active transport, with To/Cc/Bcc,
  Text or HTML body, and file attachments. Their total size is checked before anything is sent.
- Repository renamed NSP-Graph -> NSP-M365-Graph, so the folder matches the module name the way the
  shared sibling loader expects (NSP.M365.Graph -> NSP-M365-Graph). ProjectUri / LicenseUri updated.

## 0.1.0

- `Connect-NSPGraph` extracted from NSP-IntuneApps with its tests.
- `Invoke-NSPGraphRequest` / `Invoke-NSPGraphCollection` with a swappable transport
  (`MgGraphRequest`, `Fake`).
- `Register-NSPGraphAppRegistration`: generalized, REST-based tenant bootstrap for a tool's own
  delegated app registration, with plan/repair/adopt behavior, tested against an in-memory fake tenant.
- `Register-NSPGraphAppOnlyRegistration`: bootstrap for an unattended job's app, with application
  permissions and a host-created, non-exportable certificate; renews within 30 days of expiry and keeps
  the old certificate until it expires. The record includes the service principal ID.
- `Connect-NSPGraphAppOnly`: certificate sign-in that refuses anything but an app-only context.
- The client-data gate ignores Graph `@odata` annotation keys.
