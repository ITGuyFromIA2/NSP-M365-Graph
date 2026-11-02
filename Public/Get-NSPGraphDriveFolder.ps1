function Get-NSPGraphDriveFolder {
    <#
    .SYNOPSIS
        Lists a site's document libraries, or the folders inside a library folder. Read-only.
    .DESCRIPTION
        For browsing to a destination folder. -SiteId lists the site's document libraries (drives).
        -DriveId lists the folders directly under -ItemId, the library root by default. Files are
        left out. Each result carries DriveId and ItemId, the values a route needs. Uses the active
        connection; delegated callers need Sites.Read.All or Files.Read.All.
    .EXAMPLE
        Get-NSPGraphDriveFolder -SiteId $site.SiteId
    .EXAMPLE
        Get-NSPGraphDriveFolder -DriveId $library.DriveId -ItemId $folder.ItemId
    #>
    [CmdletBinding(DefaultParameterSetName = 'Folders')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Libraries')][ValidateNotNullOrEmpty()][string]$SiteId,
        [Parameter(Mandatory, ParameterSetName = 'Folders')][ValidateNotNullOrEmpty()][string]$DriveId,
        [Parameter(ParameterSetName = 'Folders')][ValidateNotNullOrEmpty()][string]$ItemId = 'root'
    )

    if ($PSCmdlet.ParameterSetName -eq 'Libraries') {
        $drives = @(Invoke-NSPGraphCollection -Uri "sites/$([uri]::EscapeDataString($SiteId))/drives?`$select=id,name,webUrl,driveType")
        return @($drives | Where-Object { $_['driveType'] -eq 'documentLibrary' } | ForEach-Object {
                [pscustomobject]@{ Kind = 'Library'; Name = [string]$_['name']; DriveId = [string]$_['id']; ItemId = 'root'; WebUrl = [string]$_['webUrl']; ChildCount = $null }
            } | Sort-Object Name)
    }
    $items = @(Invoke-NSPGraphCollection -Uri "drives/$([uri]::EscapeDataString($DriveId))/items/$([uri]::EscapeDataString($ItemId))/children?`$select=id,name,webUrl,folder&`$top=200")
    @($items | Where-Object { $null -ne $_['folder'] } | ForEach-Object {
            [pscustomobject]@{ Kind = 'Folder'; Name = [string]$_['name']; DriveId = $DriveId; ItemId = [string]$_['id']; WebUrl = [string]$_['webUrl']; ChildCount = $_['folder']['childCount'] }
        } | Sort-Object Name)
}
