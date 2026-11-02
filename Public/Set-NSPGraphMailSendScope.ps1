function Set-NSPGraphMailSendScope {
    <#
    .SYNOPSIS
        Lets an app-only registration send mail as exactly one mailbox, through Exchange RBAC for
        Applications. Plan-only by default.
    .DESCRIPTION
        Graph's Mail.Send application permission granted in Entra reaches every mailbox in the
        tenant. Exchange RBAC for Applications grants the same right limited to a management
        scope, and the two are additive: an Entra Mail.Send grant would keep tenant-wide access
        whatever the scope says. This function therefore refuses to continue (UnscopedEntraGrant)
        while the app's service principal holds Graph Mail.Send in Entra.

        Without -Execute it reports what it would do. With -Execute and ShouldProcess approval it
        creates, in order:
          - the Exchange service principal pointer for the app;
          - management scope "NSP Mail.Send <ClientId>", restricted to the mailbox's primary SMTP
            address (or updates that scope's filter when the mailbox changed);
          - an "Application Mail.Send" role assignment for the app on that scope.
        An existing Application Mail.Send assignment on any other scope, or unscoped, is reported
        as ConflictingAssignment and left alone.

        Every result that reaches a scope is verified with Test-ServicePrincipalAuthorization,
        which bypasses Exchange's permission cache; sendMail itself may take up to two hours to
        honor a new assignment.

        Needs the active Graph connection (Application.Read.All or wider) and an Exchange Online
        PowerShell session with rights to manage role assignments. -Connect imports
        ExchangeOnlineManagement and runs Connect-ExchangeOnline when no session is open.
    .EXAMPLE
        Set-NSPGraphMailSendScope -ClientId $record.ClientId -ServicePrincipalId $record.ServicePrincipalId -Mailbox 'scanner@contoso.example' -Connect
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][guid]$ClientId,
        [Parameter(Mandatory)][guid]$ServicePrincipalId,
        [Parameter(Mandatory)][ValidatePattern('^[^@\s]+@[^@\s]+$')][string]$Mailbox,
        [switch]$Execute,
        [switch]$Connect
    )

    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $roleName = 'Application Mail.Send'
    $scopeName = "NSP Mail.Send $ClientId"
    $result = [ordered]@{
        Status = $null; ClientId = [string]$ClientId; Mailbox = $Mailbox; PrimarySmtpAddress = $null
        ScopeName = $scopeName; Actions = @(); Verified = $false; Message = $null
    }

    $servicePrincipal = Invoke-NSPGraphRequest -Uri "servicePrincipals/$ServicePrincipalId`?`$select=id,appId,displayName"
    if ($servicePrincipal['appId'] -ne [string]$ClientId) {
        throw "Service principal $ServicePrincipalId does not belong to app $ClientId."
    }
    $graphSp = @(Invoke-NSPGraphCollection -Uri "servicePrincipals?`$filter=$([uri]::EscapeDataString("appId eq '$graphAppId'"))&`$select=id,appRoles")[0]
    $mailSendRoleId = @(@($graphSp['appRoles']) | Where-Object { $_ -and $_['value'] -ceq 'Mail.Send' })[0]['id']
    $entraGrant = @(Invoke-NSPGraphCollection -Uri "servicePrincipals/$ServicePrincipalId/appRoleAssignments" |
            Where-Object { $_['resourceId'] -eq $graphSp['id'] -and $_['appRoleId'] -eq $mailSendRoleId })
    if ($entraGrant.Count) {
        $result.Status = 'UnscopedEntraGrant'
        $result.Message = 'The app holds Graph Mail.Send in Entra, which reaches every mailbox and is not narrowed by an Exchange scope. Remove that admin-consented permission first; nothing was changed.'
        return [pscustomobject]$result
    }

    if ($Connect) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $open = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
        if (-not $open.Count) { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop }
    }
    if (-not (Get-Command Test-ServicePrincipalAuthorization -ErrorAction SilentlyContinue)) {
        throw 'No Exchange Online PowerShell session is open. Run Connect-ExchangeOnline, or pass -Connect.'
    }

    # Exchange cmdlets throw for a missing object even with -ErrorAction SilentlyContinue.
    $find = { param([scriptblock]$Lookup) try { & $Lookup } catch { $null } }
    $recipient = & $find { Get-Recipient -Identity $Mailbox -ErrorAction Stop }
    if (-not $recipient -or [string]$recipient.RecipientTypeDetails -notin @('UserMailbox', 'SharedMailbox')) {
        throw "'$Mailbox' is not a user or shared mailbox in this Exchange organization."
    }
    $primary = [string]$recipient.PrimarySmtpAddress
    $result.PrimarySmtpAddress = $primary
    $filter = "PrimarySmtpAddress -eq '$($primary -replace "'", "''")'"

    $actions = [System.Collections.Generic.List[string]]::new()
    $exchangeSp = & $find { Get-ServicePrincipal -Identity ([string]$ClientId) -ErrorAction Stop }
    if (-not $exchangeSp) { $actions.Add('CreateExchangeServicePrincipal') }
    $scope = & $find { Get-ManagementScope -Identity $scopeName -ErrorAction Stop }
    if (-not $scope) { $actions.Add('CreateScope') }
    elseif ([string]$scope.RecipientFilter -ne $filter) { $actions.Add('UpdateScope') }
    $assignments = if ($exchangeSp) { @(& $find { Get-ManagementRoleAssignment -RoleAssignee $exchangeSp.Identity -Role $roleName -ErrorAction Stop }) | Where-Object { $_ } } else { @() }
    $other = @($assignments | Where-Object { [string]$_.CustomResourceScope -ne $scopeName })
    if ($other.Count) {
        $result.Status = 'ConflictingAssignment'
        $result.Message = "The app already has $roleName on another or no scope: $(@($other | ForEach-Object { if ($_.CustomResourceScope) { $_.CustomResourceScope } else { '(unscoped)' } }) -join ', '). Review it in Exchange; nothing was changed."
        return [pscustomobject]$result
    }
    if (-not @($assignments).Count) { $actions.Add('CreateRoleAssignment') }
    $result.Actions = @($actions)

    if ($actions.Count -and -not $Execute) {
        $result.Status = 'PlanOnly'
        $result.Message = "Run again with -Execute to apply: $($actions -join ', ')."
        return [pscustomobject]$result
    }
    if ($actions.Count) {
        if (-not $PSCmdlet.ShouldProcess("app $ClientId", "Allow $roleName as $primary only: $($actions -join ', ')")) { return }
        if ('CreateExchangeServicePrincipal' -in $actions) {
            New-ServicePrincipal -AppId ([string]$ClientId) -ObjectId ([string]$ServicePrincipalId) -DisplayName ([string]$servicePrincipal['displayName']) -ErrorAction Stop | Out-Null
        }
        if ('CreateScope' -in $actions) { New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $filter -ErrorAction Stop | Out-Null }
        if ('UpdateScope' -in $actions) { Set-ManagementScope -Identity $scopeName -RecipientRestrictionFilter $filter -ErrorAction Stop }
        if ('CreateRoleAssignment' -in $actions) {
            New-ManagementRoleAssignment -App ([string]$ClientId) -Role $roleName -CustomResourceScope $scopeName -ErrorAction Stop | Out-Null
        }
    }

    $check = @(Test-ServicePrincipalAuthorization -Identity ([string]$ClientId) -Resource $primary -ErrorAction Stop |
            Where-Object { [string]$_.RoleName -eq $roleName })
    $result.Verified = [bool](@($check | Where-Object { $_.InScope -eq $true }).Count)
    $result.Status = if (-not $actions.Count) { 'AlreadyScoped' } else { 'Scoped' }
    $result.Message = if ($result.Verified) { "Exchange confirms $roleName is in scope for $primary only. sendMail may take up to two hours to honor a new assignment." }
        else { "Exchange did not confirm $roleName in scope for $primary. Re-check with Test-ServicePrincipalAuthorization before relying on it." }
    [pscustomobject]$result
}
