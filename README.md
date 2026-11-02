# NSP.M365.Graph

Shared Microsoft Graph plumbing for NSP M365 tools (`NSP.M365.ConditionalAccess`, and later
`NSP.M365.IntuneApps` and `NSP.M365.IntuneManager`). Supports Windows PowerShell 5.1 and PowerShell 7.

| Command | Purpose |
|---|---|
| `Connect-NSPGraph` | Connect as a delegated app, or validate the existing context, for a tool's complete scope list. Reconnects only when the session is missing a scope or belongs to a different app or tenant. |
| `Connect-NSPGraphAppOnly` | Unattended sign-in as an app with a certificate from the local store. Refuses a context that is not app-only for that client and tenant. |
| `Invoke-NSPGraphRequest` | One REST request through the active transport. Returns hashtables in Graph's JSON shape (camelCase). |
| `Invoke-NSPGraphCollection` | Every page of a collection, following `@odata.nextLink`. |
| `Register-NSPGraphAppRegistration` | Tenant bootstrap for a tool's own single-tenant, delegated public-client app: redirect URIs including the WAM broker URI, required permissions, service principal, and an admin consent grant that is only ever topped up. Plan-only without `-Execute`; never adopts an unrecorded app of the same name unless `-AdoptClientId` says so. |
| `Register-NSPGraphAppOnlyRegistration` | Tenant bootstrap for an unattended job's app: application permissions (admin-consented role assignments) and a non-exportable certificate created on the job host. Plan-only without `-Execute`; re-running within 30 days of expiry renews the certificate. |
| `Set-NSPGraphTransport` / `Get-NSPGraphTransport` | Choose the transport: `MgGraphRequest` (default) or `Fake` (a scriptblock answers each request and the calls are logged, for offline tests). |

```powershell
Import-Module .\NSP.M365.Graph.psd1
Connect-NSPGraph -Scopes 'Policy.Read.All' -Connect
$policies = Invoke-NSPGraphCollection -Uri 'identity/conditionalAccess/policies'
```

## Design

- **Token source and transport are separate.** The only token source for now is `Connect-MgGraph`,
  which keeps one context per process, so tools work with one tenant at a time. A per-tenant MSAL token
  cache with an `Invoke-RestMethod` transport can be added later, behind the same request functions,
  if concurrent multi-tenant work is adopted.
- **Cmdlet backends are the consuming module's job.** Microsoft.Graph cmdlets differ per resource
  (`Get-MgIdentityConditionalAccessPolicy`, `Get-MgGroup`, ...), so a tool offering them
  converts their output to the same REST shape inside its own backend.
- **Scopes belong to each tool.** Each consuming tool defines its permission lists and passes them to
  every connect or register call. Narrower lists at individual call sites cause repeated sign-in prompts.
- `Microsoft.Graph.Authentication` is not a `RequiredModules` entry, so the module imports on a bare host.
  `Connect-NSPGraph -Connect` installs it through NSP.Bootstrap when available.

`Register-NSPGraphAppRegistration` generalizes NSP-IntuneApps' `Register-NSPIntuneWin32AppRegistration`
over REST. Consent grants are looked up by the tool's **service principal** ID, the key Graph uses for
them. The IntuneApps original queries by the application object ID on its permission-repair path.

`Register-NSPGraphAppOnlyRegistration` keeps key material on the host. Graph replaces an app's
`keyCredentials` as a whole and never returns keys, so every valid `CN=<AppName>` certificate in
the local store is re-sent on each update. That makes one job host per tenant. The record
(`<RecordDirectory>\<TenantId>.json`) holds the client ID, service principal ID, and certificate
thumbprint; it never holds a key.

### Sending as one mailbox

Mail.Send granted in Entra reaches every mailbox in the tenant. To let an unattended job send as a
single mailbox:

1. Find the sender with `Find-NSPGraphMailbox -Search scan` (read-only; `-Address` for an exact
   lookup). `MailboxHint` comes from sign-in and license state, not from Exchange.
2. Register the app with `-ApplicationPermissions @()`, so it has no Graph permission at all.
3. Run `Set-NSPGraphMailSendScope -ClientId ... -ServicePrincipalId ... -Mailbox ... -Connect`. It plans
   first; add `-Execute` to create the Exchange service principal, a management scope for that
   mailbox's primary address, and an `Application Mail.Send` assignment on that scope.

Exchange RBAC grants add to Entra grants, so the function refuses while the app still holds Graph
Mail.Send in Entra. It also refuses when an Application Mail.Send assignment already exists on
another scope or none. Results are checked with `Test-ServicePrincipalAuthorization`. Exchange can
take up to two hours before sendMail honors a new assignment. This function needs
ExchangeOnlineManagement; nothing else in the module does.

## Validation

```powershell
.\tools\Test-Repo.ps1
```

## Certificate authentication context

`Connect-NSPGraphAppOnly` accepts optional `-ContextScope Process` or `-ContextScope CurrentUser`.
Omitting it preserves the SDK's existing scope-selection behavior; this module assigns no default.
An explicit choice is forwarded and the returned context must match that choice as well as the
requested tenant, application and AppOnly authentication type. A mismatch fails without retry.

Process confines authentication to the current PowerShell session. CurrentUser allows persistence
across sessions. This choice does not grant permissions or make a certificate/private key accessible
to another account. The SDK must already be installed; the app-only connector does not install it.
See Microsoft's [authentication commands](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands).

```powershell
Connect-NSPGraphAppOnly -ClientId $record.ClientId -TenantId $record.TenantId `
    -CertificateThumbprint $record.CertificateThumbprint -ContextScope Process
```

The isolated `tools/Test-AppOnlyContextScope.ps1` regression shadows SDK commands locally and uses
synthetic identities. It can run directly under PowerShell 5.1 or 7 without the SDK, certificates,
tenant access or test-framework installation, and is included in `tools/Test-Repo.ps1`.