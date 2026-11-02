function Find-NSPGraphSite {
    <#
    .SYNOPSIS
        Finds SharePoint sites by name, or resolves a sharing or browser link to its site, library and item. Read-only.
    .DESCRIPTION
        For choosing a destination without copying IDs by hand. Uses the active connection; delegated
        callers need Sites.Read.All.

        -Search returns sites whose name or description matches (Graph sites?search=).
        -Url accepts a SharePoint link to a folder or file, as copied from the browser address bar
        or Copy link. It is resolved through the shares API, and the result carries the site,
        library (drive) and item IDs. A link the signed-in user cannot open does not resolve.
        -DriveId returns the site that holds a library, for example to grant an app on the site
        behind a configured route.
    .EXAMPLE
        Find-NSPGraphSite -Search 'Scans'
    .EXAMPLE
        Find-NSPGraphSite -Url 'https://contoso.sharepoint.com/sites/Scans/Shared%20Documents/Inbox'
    #>
    [CmdletBinding(DefaultParameterSetName = 'Search')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Search')][ValidateNotNullOrEmpty()][string]$Search,
        [Parameter(Mandatory, ParameterSetName = 'Url')][ValidatePattern('^https://[^/]+\.sharepoint\.com/')][string]$Url,
        [Parameter(Mandatory, ParameterSetName = 'Drive')][ValidateNotNullOrEmpty()][string]$DriveId
    )

    if ($PSCmdlet.ParameterSetName -eq 'Drive') {
        # A library root has no parent, so the site comes from the drive's SharePoint IDs, combined
        # into Graph's site ID form: hostname,site collection ID,web ID.
        $drive = Invoke-NSPGraphRequest -Uri "drives/$([uri]::EscapeDataString($DriveId))?`$select=id,name,webUrl,sharePointIds"
        $ids = $drive['sharePointIds']
        if (-not $ids -or -not $ids['siteUrl'] -or -not $ids['siteId'] -or -not $ids['webId']) { throw "Library $DriveId did not return its SharePoint site IDs; it may not be a SharePoint document library." }
        return [pscustomobject]@{
            SiteId = '{0},{1},{2}' -f ([uri][string]$ids['siteUrl']).Host, $ids['siteId'], $ids['webId']
            DriveId = $DriveId; ItemId = 'root'; Name = [string]$drive['name']; WebUrl = [string]$drive['webUrl']; IsFolder = $true
        }
    }

    if ($PSCmdlet.ParameterSetName -eq 'Search') {
        $sites = @(Invoke-NSPGraphCollection -Uri "sites?search=$([uri]::EscapeDataString($Search.Trim()))&`$select=id,displayName,name,webUrl")
        return @($sites | ForEach-Object {
                [pscustomobject]@{ SiteId = [string]$_['id']; DisplayName = [string]$_['displayName']; WebUrl = [string]$_['webUrl'] }
            } | Sort-Object DisplayName)
    }
    # Shares API: "u!" + unpadded base64url of the link.
    $token = 'u!' + ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Url.Trim())).TrimEnd('=').Replace('/', '_').Replace('+', '-'))
    $item = Invoke-NSPGraphRequest -Uri "shares/$token/driveItem?`$select=id,name,webUrl,folder,file,parentReference"
    $parent = $item['parentReference']
    [pscustomobject]@{
        SiteId = [string]$parent['siteId']; DriveId = [string]$parent['driveId']; ItemId = [string]$item['id']
        Name = [string]$item['name']; WebUrl = [string]$item['webUrl']; IsFolder = $null -ne $item['folder']
    }
}
