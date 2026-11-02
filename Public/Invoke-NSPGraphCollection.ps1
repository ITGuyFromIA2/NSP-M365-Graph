function Invoke-NSPGraphCollection {
    <#
    .SYNOPSIS
        Reads every page of a Graph collection and returns the items.
    .DESCRIPTION
        Follows @odata.nextLink until the collection is exhausted. Each page goes through
        Invoke-NSPGraphRequest, so the active transport applies. -Headers is sent with every page,
        as $search and advanced queries require.
    .EXAMPLE
        Invoke-NSPGraphCollection -Uri 'identity/conditionalAccess/namedLocations'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0',
        [hashtable]$Headers
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $requestArgs = @{ Method = 'GET'; ApiVersion = $ApiVersion }
    if ($Headers) { $requestArgs.Headers = $Headers }
    while ($next) {
        $page = Invoke-NSPGraphRequest -Uri $next @requestArgs
        foreach ($item in @($page['value'])) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $next = $page['@odata.nextLink']
    }
    $items.ToArray()
}
