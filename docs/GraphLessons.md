# Graph lessons

Read this before writing Graph code in any NSP tool. Each lesson below cost a consumer a failed run
or a wrong answer. The fix usually lives in that one consumer, so the next tool will not inherit it
from this module unless you apply it yourself.

Each lesson gives the rule, what happened, and where it is recorded. **Observed** means it happened
in a live run. **By design** means it is how Graph, the SDK or this module is built. **Reasoned**
means the rule follows from code or from another lesson, and has not been seen live in that form.
Sources marked *(private)* are in repositories that are not public; the lesson stands without them.

To add a lesson, name its source, and keep client names, tenant IDs and raw errors out. This file
is public.

## Request bodies

### Lists unroll into scalars or null

Observed. Assign a list directly. In `$x = if (...) { @(...) }`, PowerShell unrolls the result, so
`@('All')` goes out as `"All"` and `@()` as `null`. Graph answers 400 with error 1007, "does not
match the schema". A list of one hashtable unrolls the same way and serializes as an object.

The in-memory fake accepted both shapes, so the tests passed. The first live write caught it.

Keep bodies in the fake as serialized JSON, and test that each list field is a JSON array.

Sources: NSP-M365-ConditionalAccess *(private)* `PLAN.md` "First live execution", `CLAUDE.md`
"PowerShell pitfalls", `Tests/Run.Tests.ps1`; `Public/Send-NSPGraphMail.ps1` here.

### Property order can matter

Observed. Use `[ordered]` for a body that carries an `@microsoft.graph.*` annotation, and put the
annotation first. An upload-session item with `name` before `@microsoft.graph.conflictBehavior`
returned 400; annotation-first succeeded. A plain hashtable serialized in a different order across
fresh processes, so this failed intermittently.

Sources: NSP-ScanBridge `Private/Send-NSPScanBridgeUploadSession.ps1`; Coordination *(private)*
`coordination/notes/2026-10-30_from-scanbridge_to-ptmaster_live-auth-review.md`.

## Reading after writing

### Reads lag writes

Observed. Do not conclude anything from a single read straight after a write. A read-back right
after a Conditional Access PATCH returned the old policy, and a re-inventory minutes later showed
the change had applied. ConditionalAccess now re-reads after 2, 4, 8, 15 and 30 seconds.

Report "still reads as before" separately from "reads differently". The first may be lag, and the
second means Graph stored something other than what you sent.

Reasoned: a hard-match read-back of `onPremisesImmutableId` needs the same treatment. The lag was
observed on policies only; user objects have not been measured.

Sources: NSP-M365-ConditionalAccess *(private)* `PLAN.md` "First live execution",
`ConfirmApplied` in `Private/Invoke-NSPConditionalAccessRunStage.ps1`.

### Exchange scopes take up to two hours

Observed. A `sendMail` through a new Exchange RBAC for Applications assignment returned 403 at
first. It cleared about 1 hour 38 minutes after setup. Treat a 403 in that window as propagation,
and do not repeat the setup.

`Test-ServicePrincipalAuthorization` bypasses Exchange's permission cache, so it can confirm the
assignment before `sendMail` honors it. A pass there does not prove mail will send yet.

Sources: `Public/Set-NSPGraphMailSendScope.ps1` here; NSP-Toolkit
`Private/ScanBridgeAlertSetup.ps1`; Coordination *(private)*
`coordination/notes/2026-11-03_from-scanbridge_to-ptmaster_alert-test-passed.md`.

## Paging and large reads

### A partial read is not a result

Observed. Record whether a collection finished, and keep the `@odata.nextLink` when it did not. A
page of a sign-in log export failed on the SDK's limit, "HttpClient.Timeout of 300 seconds",
after two pages of 1,000 records had been saved. The run stayed correct because that stream was
marked incomplete with its nextLink.

The cause of the wait is undetermined: nothing captured separates SDK retries, service latency,
payload size and network. Do not call it throttling. A longer timeout only prolongs the silence.
Use `$select`, a sensible `$top`, and narrower time windows, and show page, count and elapsed time
so a blocked request is visible.

`Invoke-NSPGraphCollection` returns everything or throws: when a page fails, the pages already read
are discarded. That is safe, and it is the wrong tool for a read you need to resume. Page with
`Invoke-NSPGraphRequest` there.

Sources: NSP-M365-Investigation `Private/HuntHelpers.ps1` (`Get-NSPHuntCollection`), and its
collection helpers (uncommitted at the time of writing); Coordination *(private)*
`workstreams/m365-investigation/COLLECTION-PERFORMANCE-NOTES.md`.

### Guard the paging loop

Reasoned. In a loop you write yourself, stop on a repeated nextLink, and treat reaching a page cap
as incomplete, never as the end of the data. Do not lean on the value check alone: `@($page['value'])`
yields one `$null` item when the key is missing, so a response with no `value` looks like a page
with nothing in it.

Sources: `Public/Invoke-NSPGraphCollection.ps1` and `Public/Register-NSPGraphAppRegistration.ps1`
(`Get-NSPGraphItem`) here; NSP-M365-Investigation collection helpers (uncommitted at the time of
writing).

### `$search` and `$count` need a header

By design. Send `ConsistencyLevel: eventual` with them. `Invoke-NSPGraphRequest` and
`Invoke-NSPGraphCollection` take `-Headers` and send it with every page. `$search` matches the
start of words and is eventually consistent, so it is for picking, not for proving absence.

Source: `Public/Find-NSPGraphMailbox.ps1` here.

### Retry only what is safe to repeat

By design. Neither collection function retries on its own, though the SDK may retry inside one
call. Retry a GET on 429 and transient 5xx responses with growing waits. For a write, retry only
while the result cannot be ambiguous: ScanBridge retries before an upload session exists, and
never after, because by then the outcome is uncertain.

Sources: NSP-M365-Investigation `Private/HuntHelpers.ps1` (`Invoke-NSPGraphRetry`); NSP-ScanBridge
`Private/Test-NSPScanBridgeRetrySafe.ps1`.

## Sessions and sign-in

### One tenant per process

By design. `Connect-MgGraph` and `Invoke-MgGraphRequest` serve one context at a time.
Concurrent tenants would need MSAL plus REST, which risks clashing with the copy of MSAL bundled in
Microsoft.Graph.Authentication; that has not been tested. Anything holding tenant-bound state must
compare the session's tenant with its own and refuse a mismatch.

Sources: NSP-M365-ConditionalAccess *(private)* `PLAN.md` "Graph layer"; "Design" in this
repository's README.

### Check who the session is, not only that one exists

Observed. A leftover session from the SDK's default app can hold the right scopes and still be the
wrong identity. Compare tenant, client ID, auth type and scopes before work starts, and name the
reconnect command in the error.

Sources: NSP-M365-ConditionalAccess *(private)* `Private/Assert-NSPConditionalAccessSession.ps1`;
NSP-ScanBridge `Private/Connect-NSPScanBridgeContext.ps1`.

### Request scopes from one place

Observed. Define a tool's scope list once and use it at every call site. A narrower list at one
call site makes whichever action runs first set up a narrow token, and the next action prompts
again.

Sources: NSP-M365-ConditionalAccess *(private)*
`Private/Get-NSPConditionalAccessGraphScopes.ps1`, citing the NSP-IntuneApps lesson.

### App-only certificate context

Observed on one workstation. With SDK 2.40.0, certificate sign-in failed with `ContextScope
CurrentUser` and worked with `Process`. The cause is undetermined, and it was not a tenant
permission problem. `Connect-NSPGraphAppOnly` takes an optional `-ContextScope` and validates the
returned scope.

**Do not change the default.** The owner rejected that: omitting the parameter must keep the SDK's
behavior. A consumer that needs `Process` asks for it, as ScanBridge does.

Sources: "Certificate authentication context" in this repository's README; Coordination *(private)*
`coordination/notes/2026-10-30_from-scanbridge_to-ptmaster_live-auth-review.md`.

### Sign-in quirks not yet explained

Observed, cause unknown; note a repeat rather than designing around these.

- Windows PowerShell 5.1 did not discover SDK 2.40.0 by name. Loading its manifest by path worked.
- Device-code authentication returned a null-reference error once; browser sign-in succeeded.
- An extra sign-in prompt appeared at the first read after connecting, from the SDK's own token
  acquisition.
- Over GDAP, a delegated sign-in to a client tenant needs the "Microsoft Graph Command Line Tools"
  app consented there, or it fails with AADSTS650051 or AADSTS65001.

Sources: the live-auth note above; NSP-M365-ConditionalAccess *(private)* `PLAN.md` "First live
execution"; NSP-M365-Investigation `README.md`.

### Private keys and elevation

Observed. A non-elevated session reading a computer-store certificate key fails with "Keyset does
not exist" (0x80090016). That key is readable only by administrators and the account granted
access, so it says nothing about the app or the tenant. Run the test elevated, or as that account.

Sources: NSP-Toolkit `Private/ScanBridgeAlertSetup.ps1`; Coordination *(private)*
`coordination/notes/2026-11-03_from-scanbridge_to-ptmaster_alert-test-passed.md`.

## Loading this module

### Resolve the version inside your own module

Observed. ScanBridge's worker imported an older installed copy inside its module scope even with a
newer copy imported globally. Import inside your module's scope with `-MinimumVersion`, which also
skips an older copy earlier on `PSModulePath`, and call the command through the imported module.

Sources: NSP-ScanBridge `Private/Connect-NSPScanBridgeContext.ps1`; the live-auth note above.

### A long-lived shell keeps old code

Observed twice. A session with the module already loaded keeps running the old functions. In one
case a re-run lacked the read-back fix and reported a false `NeedsReview`; in another an expected
new command was missing. A reused module from the same path is not proof of the current code.
Start a fresh process after an update, or re-import with `-Force`.

Sources: NSP-M365-ConditionalAccess *(private)* `PLAN.md` "First live execution"; Coordination
*(private)* `workstreams/m365-investigation/CONTINUATION.md`.

### Connecting can install

By design. `Connect-NSPGraph -Connect` installs a missing SDK through NSP.Bootstrap when available,
and this module does not check `NSP_NO_AUTOINSTALL`. A tool that must not install checks its
dependencies first. `Connect-NSPGraphAppOnly` never installs.

Sources: "Design" in this repository's README; Coordination *(private)*
`workstreams/m365-investigation/CONTINUATION.md`.

## Identity and directory data

### Names are hints, never evidence

Observed. Breakglass accounts were not findable by name; they came from who the existing policies
exclude. The real directory-sync account was not the one named like one: an operator marked the
look-alike as the service account, and only preflight noticed the real one was not excluded from
a block policy. Use role membership and object references.
Microsoft-managed policies carry `templateId`; leave those out of matching.

Reasoned: a pairing scorer fails the same way, so a name match can suggest a pair and never settle
one.

Sources: NSP-M365-ConditionalAccess *(private)* `PLAN.md` "Findings from the first dev-tenant run"
and "Second dev-tenant run".

### Keys that are easy to get wrong

Observed. Consent grants are keyed by the **service principal** object ID, not the application's;
NSP-IntuneApps' original looks up the application object ID on its repair path. Identify people and
apps by object ID, and treat a UPN as a label.

Sources: `Public/Register-NSPGraphAppRegistration.ps1` and the README here; NSP-M365-ConditionalAccess
*(private)* `PLAN.md` "Open items".

### `keyCredentials` is replaced whole

By API design. Graph replaces an app's `keyCredentials` as a whole and never returns keys, so an
update must re-send every certificate to keep. That is why each tenant has one job host for an
app-only registration.

Source: this repository's README.

## SharePoint and drives

### A library root has no parent

Observed. Reading the site from a drive root's `parentReference` returned nothing, because a root
has no parent. Read `drive.sharePointIds` and build Graph's `hostname,siteId,webId` form.

Source: `Public/Find-NSPGraphSite.ps1` here, commit 62bf77c.

### Sites.Selected reaches nothing until the site names the app

By design. The app's permission only makes a site grant possible; each site must then name the app.
Upgrade a weaker grant in place, and leave other apps' permissions alone.

Sources: NSP-Toolkit `Private/ScanBridgeRouteSetup.ps1`; `Public/Grant-NSPGraphSitePermission.ps1`
here.

## Testing

### The fake accepts what Graph might not

Observed. The list-shape and library-root defects both passed offline tests first. A live finding
goes into the tests as well as the code, and a fake that stores bodies as serialized JSON catches more than one
that keeps the live object. This module's `Fake` transport passes `$Body` to the responder
unserialized, and logs `-Headers` without passing them, so assert headers from `CallLog`.

Sources: NSP-M365-ConditionalAccess *(private)* `CLAUDE.md` "Testing";
`Public/Invoke-NSPGraphRequest.ps1` here.
