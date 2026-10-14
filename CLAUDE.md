# CLAUDE.md

## What this is

`NSP.M365.Graph`: shared Graph connection, paging, and request transport for the NSP M365 tools.
Toolkit B4 in `C:\GitRepo\PoSHRepo\ClaudeStuff\04_Toolkit_Candidates.md`. The first consumer is
`NSP-PoSHToolkits\NSP-M365-ConditionalAccess` (see its PLAN.md, "Graph layer").

## Hard rules

- Windows PowerShell 5.1 is the floor. The module must import on a bare host; Graph SDK modules are not
  `RequiredModules`.
- Nothing domain-specific lives here. No Conditional Access, Intune, or group logic, and no scope lists.
- Registration is plan-only without `-Execute`. App-only registrations keep private keys on the host:
  certificates are non-exportable, and records hold thumbprints, never keys or secrets.
- Tests never touch a real certificate store or tenant: certificate helpers are private wrappers
  that tests mock, and Graph goes through the `Fake` transport.
- The canonical data shape is Graph REST JSON as hashtables (camelCase).
- Never abbreviate Conditional Access as "CA" (NSP-CAManager owns that noun space).
- One function per file. `FunctionsToExport` matches `Public\`. Exported functions need `.SYNOPSIS` and `.EXAMPLE`.

## Testing

Pester 5+. Run `.\tools\Test-Repo.ps1`.

## Publishing (not adopted yet - do this on next use)

This repo isn't on the shared NSP release routine yet. Before its next release, adopt it the way
the published toolkit repos have (NSP-Console, NSP-Toolkit, NSP-ActiveDirectory, NSP-Bootstrap):

- Add `tools\Publish-ToGallery.ps1`: the same thin wrapper those repos carry, which imports
  NSP.RepoTools and runs `Publish-NSPModule` on this repo. Copy it as-is; all the logic stays in
  NSP.RepoTools.
- Dry run first: `.\tools\Publish-ToGallery.ps1 -WhatIf` runs every check (manifest, changelog,
  clean tree, Gallery version, client references across history, `tools\Test-Repo.ps1`, staged
  5.1 import) and publishes nothing.
- At release, change the CHANGELOG heading from `## <version> (unreleased)` to `## <version>`;
  the routine refuses an unreleased heading.
- Needs a GitHub remote first if the repo has none (the manifest's ProjectUri should point at it).