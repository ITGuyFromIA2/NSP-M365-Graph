function New-NSPGraphDriveFolder {
    <#
    .SYNOPSIS
        Creates one folder in a SharePoint library, failing rather than renaming when the name is taken.
    .DESCRIPTION
        Creates -Name under -ParentId (the library root by default) and returns its DriveId and
        ItemId. The conflict behavior is fail, so an existing folder is never replaced or given a
        "Name 1" twin. Uses the active connection; delegated callers need Sites.ReadWrite.All or
        Files.ReadWrite.All.
    .EXAMPLE
        New-NSPGraphDriveFolder -DriveId $library.DriveId -Name 'ScanBridge Test'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DriveId,
        [ValidateNotNullOrEmpty()][string]$ParentId = 'root',
        [Parameter(Mandatory)][ValidatePattern('^[^"*:<>?/\\|#%]{1,250}$')][string]$Name
    )

    if (-not $PSCmdlet.ShouldProcess("drive $DriveId", "Create folder '$Name'")) { return }
    $folder = Invoke-NSPGraphRequest -Method POST -Uri "drives/$([uri]::EscapeDataString($DriveId))/items/$([uri]::EscapeDataString($ParentId))/children" -Body @{
        name = $Name.Trim(); folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail'
    }
    [pscustomobject]@{ Kind = 'Folder'; Name = [string]$folder['name']; DriveId = $DriveId; ItemId = [string]$folder['id']; WebUrl = [string]$folder['webUrl']; ChildCount = 0 }
}
