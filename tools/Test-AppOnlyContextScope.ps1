<#
.SYNOPSIS
    Offline contract checks for certificate context-scope selection. No SDK or tenant access.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
& {
    # Scope the fake SDK commands to this invocation; do not replace the caller's commands.
    function Import-Module { [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[CmdletBinding()]param([string]$Name) if ($Name -ne 'Microsoft.Graph.Authentication') { throw 'Unexpected import.' }; $state.Imports++ }
    function Connect-MgGraph {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[CmdletBinding()]
        param($ClientId,$TenantId,$CertificateThumbprint,[switch]$NoWelcome,[ValidateSet('Process','CurrentUser')]$ContextScope)
        $state.Calls++
        $state.Captured=@{}; foreach ($key in $PSBoundParameters.Keys) { $state.Captured[$key]=$PSBoundParameters[$key] }
        if ($state.ConnectFails) { throw 'Synthetic connection failure.' }
        $selected=if ($PSBoundParameters.ContainsKey('ContextScope')) { $ContextScope } else { $state.DefaultScope }
        $state.Context=[pscustomobject]@{ ClientId=$ClientId; TenantId=$TenantId; AuthType='AppOnly'; ContextScope=$selected }
        if ($state.OverrideField) { $state.Context.($state.OverrideField)=$state.OverrideValue }
    }
    function Get-MgContext { [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline contract intentionally shadows commands and accepts their signature; no live changes.')]
[CmdletBinding()]param() if ($state.NullContext) { return $null }; $state.Context }
    function Check { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message }; $state.Checks++ }
    function Reset-Fake {
        $state.Calls=0; $state.Imports=0; $state.Captured=@{}; $state.Context=$null
        $state.OverrideField=$null; $state.OverrideValue=$null; $state.NullContext=$false
        $state.ConnectFails=$false; $state.DefaultScope='CurrentUser'
    }
    function Expect-Failure {
        param([scriptblock]$Action)
        $failed=$false; try { & $Action | Out-Null } catch { $failed=$true }
        Check $failed 'Expected fail-closed rejection.'
    }
    $state=@{ Checks=0 }
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'Public/Connect-NSPGraphAppOnly.ps1')
    $identity=@{ ClientId='synthetic-app'; TenantId='synthetic-tenant'; CertificateThumbprint='SYNTHETIC' }
    foreach ($default in 'CurrentUser','Process') {
        Reset-Fake; $state.DefaultScope=$default
        $result=Connect-NSPGraphAppOnly @identity
        Check ($result.ContextScope -eq $default) 'Omitted scope changed SDK selection.'
        Check (-not $state.Captured.ContainsKey('ContextScope')) 'Omitted scope was forwarded.'
        Check ($state.Calls -eq 1 -and $state.Imports -eq 1) 'Unexpected connection count.'
    }
    foreach ($scope in 'CurrentUser','Process') {
        Reset-Fake
        $result=Connect-NSPGraphAppOnly @identity -ContextScope $scope
        Check ($state.Captured.ContextScope -eq $scope -and $result.ContextScope -eq $scope) 'Explicit scope not forwarded/returned.'
        Check ($state.Captured.ClientId -eq $identity.ClientId -and $state.Captured.TenantId -eq $identity.TenantId -and $state.Captured.CertificateThumbprint -eq $identity.CertificateThumbprint) 'Identity forwarding changed.'
        Check ([bool]$state.Captured.NoWelcome -and $state.Captured.ErrorAction -eq 'Stop') 'SDK control parameters changed.'
        foreach ($wrong in @('other',$null)) {
            Reset-Fake; $state.OverrideField='ContextScope'; $state.OverrideValue=$wrong
            Expect-Failure { Connect-NSPGraphAppOnly @identity -ContextScope $scope }
        }
    }
    foreach ($field in 'ClientId','TenantId','AuthType') {
        foreach ($explicit in $false,$true) {
            Reset-Fake; $state.OverrideField=$field; $state.OverrideValue='wrong'
            $argsForConnection=@{}; if ($explicit) { $argsForConnection.ContextScope='Process' }
            Expect-Failure { Connect-NSPGraphAppOnly @identity @argsForConnection }
        }
    }
    Reset-Fake; $state.NullContext=$true
    Expect-Failure { Connect-NSPGraphAppOnly @identity -ContextScope Process }
    Reset-Fake; $state.ConnectFails=$true
    Expect-Failure { Connect-NSPGraphAppOnly @identity -ContextScope Process }
    Check ($state.Calls -eq 1) 'Connection failure retried unexpectedly.'
    Reset-Fake
    Expect-Failure { Connect-NSPGraphAppOnly @identity -ContextScope Invalid }
    Check ($state.Calls -eq 0 -and $state.Imports -eq 0) 'Invalid scope reached SDK/import.'
    Write-Host ('Passed: {0} app-only context-scope checks in PowerShell {1}; SDK and network mocked.' -f $state.Checks,$PSVersionTable.PSVersion)
}
