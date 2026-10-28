# Changelog

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
