function Connect-NSPGraphAppOnly {
    <#
    .SYNOPSIS
        Unattended sign-in: connects as an app with a certificate from the local store, and returns the context.
    .DESCRIPTION
        For scheduled jobs only. The certificate is referenced by thumbprint and must be in the
        store of the account the job runs as. Refuses a context that is not app-only for that
        client and tenant. ContextScope is optional: when omitted, the SDK retains its existing
        scope selection. An explicit Process or CurrentUser choice is forwarded and verified.
    .PARAMETER ContextScope
        Process limits authentication to this PowerShell session; CurrentUser permits persistence
        across sessions. Omitting this parameter preserves the SDK's scope-selection behavior.
    .EXAMPLE
        Connect-NSPGraphAppOnly -ClientId $record.ClientId -TenantId $record.TenantId -CertificateThumbprint $record.CertificateThumbprint
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$CertificateThumbprint,
        [ValidateSet('Process','CurrentUser')][string]$ContextScope
    )

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $connectArgs = @{ ClientId = $ClientId; TenantId = $TenantId; CertificateThumbprint = $CertificateThumbprint; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('ContextScope')) { $connectArgs.ContextScope = $ContextScope }
    Connect-MgGraph @connectArgs | Out-Null
    $context = Get-MgContext
    if (-not $context -or [string]$context.AuthType -ne 'AppOnly' -or $context.ClientId -ne $ClientId -or $context.TenantId -ne $TenantId) {
        throw "App-only sign-in did not produce the expected context (client $ClientId, tenant $TenantId)."
    }
    if ($PSBoundParameters.ContainsKey('ContextScope') -and [string]$context.ContextScope -ne $ContextScope) {
        throw 'App-only sign-in did not produce the requested context scope.'
    }
    $context
}
