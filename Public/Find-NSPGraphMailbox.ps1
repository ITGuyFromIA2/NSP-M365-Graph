function Find-NSPGraphMailbox {
    <#
    .SYNOPSIS
        Finds mail-enabled member accounts by a search term or an exact address, with a hint of
        what kind of mailbox each one probably is. Read-only.
    .DESCRIPTION
        For picking a sender mailbox (for example a copier's "scanner@" mailbox) without typing
        the address from memory. Uses the active connection; delegated callers need User.Read.All.

        -Search matches the start of any word in the display name, mail address or user principal
        name (Graph $search), so 'scan' finds "Scanner", scanner@ and front-desk-scan@. -Address
        resolves one exact mail, user principal name or SMTP proxy address.

        Accounts without a mail address and guests are left out: they have no mailbox in this
        tenant to send from. MailboxHint comes from Entra sign-in and license state, not from
        Exchange, so treat it as a hint:
          - LikelyShared: sign-in disabled and unlicensed, the usual shape of a shared mailbox.
          - LicensedUser: an enabled, licensed account; a person or a device that signs in.
          - Unlicensed: enabled but unlicensed; may have no mailbox at all.
          - DisabledLicensed: sign-in disabled but still licensed.
    .EXAMPLE
        Find-NSPGraphMailbox -Search 'scan'
    .EXAMPLE
        Find-NSPGraphMailbox -Address 'scanner@contoso.example'
    #>
    [CmdletBinding(DefaultParameterSetName = 'Search')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Search')][ValidatePattern('^[^"\\]+$')][string]$Search,
        [Parameter(Mandatory, ParameterSetName = 'Address')][ValidatePattern('^[^@\s]+@[^@\s]+$')][string]$Address,
        [ValidateRange(1, 999)][int]$MaxResults = 50
    )

    $select = 'id,displayName,mail,userPrincipalName,accountEnabled,assignedLicenses,userType'
    $headers = @{ ConsistencyLevel = 'eventual' }
    if ($PSCmdlet.ParameterSetName -eq 'Search') {
        $term = $Search.Trim()
        $query = (@('displayName', 'mail', 'userPrincipalName') | ForEach-Object { '"{0}:{1}"' -f $_, $term }) -join ' OR '
        $uri = "users?`$search=$([uri]::EscapeDataString($query))&`$select=$select&`$top=$([Math]::Min($MaxResults, 100))&`$count=true"
    } else {
        $value = $Address.Trim() -replace "'", "''"
        $filter = "mail eq '$value' or userPrincipalName eq '$value' or proxyAddresses/any(p:p eq 'smtp:$value')"
        $uri = "users?`$filter=$([uri]::EscapeDataString($filter))&`$select=$select&`$count=true"
    }

    $users = @(Invoke-NSPGraphCollection -Uri $uri -Headers $headers)
    $found = foreach ($user in $users) {
        if (-not $user['mail'] -or $user['userType'] -eq 'Guest') { continue }
        $enabled = [bool]$user['accountEnabled']
        $licensed = @($user['assignedLicenses'] | Where-Object { $null -ne $_ }).Count -gt 0
        $hint = if (-not $enabled -and -not $licensed) { 'LikelyShared' }
            elseif ($enabled -and $licensed) { 'LicensedUser' }
            elseif ($enabled) { 'Unlicensed' }
            else { 'DisabledLicensed' }
        [pscustomobject]@{
            Mail = [string]$user['mail']
            DisplayName = [string]$user['displayName']
            UserPrincipalName = [string]$user['userPrincipalName']
            AccountEnabled = $enabled
            Licensed = $licensed
            MailboxHint = $hint
            Id = [string]$user['id']
        }
    }
    @($found | Sort-Object @{ Expression = { $_.MailboxHint -ne 'LikelyShared' } }, DisplayName | Select-Object -First $MaxResults)
}
