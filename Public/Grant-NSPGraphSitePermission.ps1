function Grant-NSPGraphSitePermission {
    <#
    .SYNOPSIS
        Gives an app with Sites.Selected access to one SharePoint site. Plan-only by default.
    .DESCRIPTION
        Sites.Selected consents an app to no site at all until a site permission names it. This
        adds that permission for -ClientId on -SiteId with -Role (write by default; read for
        read-only jobs). Without -Execute it reports the plan; with -Execute and ShouldProcess
        approval it creates the permission, or upgrades a weaker existing one. An existing equal
        or stronger role is reported as AlreadyGranted and left alone. Other apps' permissions are
        never touched.

        Uses the active connection. Creating site permissions delegated needs a SharePoint or
        global administrator with Sites.FullControl.All.
    .EXAMPLE
        Grant-NSPGraphSitePermission -SiteId $site.SiteId -ClientId $record.ClientId -DisplayName 'NSP-ScanBridge-Ingest' -Execute
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SiteId,
        [Parameter(Mandatory)][guid]$ClientId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DisplayName,
        [ValidateSet('read', 'write')][string]$Role = 'write',
        [switch]$Execute
    )

    $rank = @{ read = 1; write = 2; manage = 3; fullcontrol = 4 }
    $sitePath = "sites/$([uri]::EscapeDataString($SiteId))/permissions"
    $result = [ordered]@{ Status = $null; SiteId = $SiteId; ClientId = [string]$ClientId; Role = $Role; ExistingRoles = @(); PermissionId = $null; Message = $null }
    $mine = @(Invoke-NSPGraphCollection -Uri $sitePath | Where-Object {
            @(@($_['grantedToIdentitiesV2']) + @($_['grantedToIdentities']) | Where-Object { $_ -and $_['application'] -and $_['application']['id'] -eq [string]$ClientId }).Count
        })
    $existing = @($mine | ForEach-Object { @($_['roles']) } | Where-Object { $_ })
    $result.ExistingRoles = $existing
    $best = @($existing | ForEach-Object { $rank[[string]$_] } | Measure-Object -Maximum).Maximum
    if ($best -ge $rank[$Role]) {
        $result.Status = 'AlreadyGranted'; $result.PermissionId = [string]$mine[0]['id']
        $result.Message = "The app already has $($existing -join ', ') on this site."
        return [pscustomobject]$result
    }
    if (-not $Execute) {
        $result.Status = 'PlanOnly'
        $result.Message = $(if ($mine.Count) { "Run again with -Execute to raise the app's site role to $Role." } else { "Run again with -Execute to grant the app $Role on this site." })
        return [pscustomobject]$result
    }
    if (-not $PSCmdlet.ShouldProcess("site $SiteId", "Grant $Role to app $DisplayName ($ClientId)")) { return }
    if ($mine.Count) {
        Invoke-NSPGraphRequest -Method PATCH -Uri "$sitePath/$($mine[0]['id'])" -Body @{ roles = @($Role) } | Out-Null
        $result.PermissionId = [string]$mine[0]['id']; $result.Status = 'Upgraded'
    } else {
        $created = Invoke-NSPGraphRequest -Method POST -Uri $sitePath -Body @{
            roles = @($Role); grantedToIdentities = @(@{ application = @{ id = [string]$ClientId; displayName = $DisplayName } })
        }
        $result.PermissionId = [string]$created['id']; $result.Status = 'Granted'
    }
    $result.Message = "The app has $Role on this site."
    [pscustomobject]$result
}
