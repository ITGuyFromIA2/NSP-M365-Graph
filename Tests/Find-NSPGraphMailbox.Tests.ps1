BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop
    $script:users = @(
        @{ id = 'u1'; displayName = 'Front Desk Scanner'; mail = 'scanner@contoso.example'; userPrincipalName = 'scanner@contoso.example'; accountEnabled = $false; assignedLicenses = @(); userType = 'Member' }
        @{ id = 'u2'; displayName = 'Scan Operator'; mail = 'scan.op@contoso.example'; userPrincipalName = 'scan.op@contoso.example'; accountEnabled = $true; assignedLicenses = @(@{ skuId = 'sku' }); userType = 'Member' }
        @{ id = 'u3'; displayName = 'Scanner Service'; mail = $null; userPrincipalName = 'svc-scan@contoso.example'; accountEnabled = $true; assignedLicenses = @(); userType = 'Member' }
        @{ id = 'u4'; displayName = 'Scanner Guest'; mail = 'guest@fabrikam.example'; userPrincipalName = 'guest_fabrikam.example#EXT#@contoso.example'; accountEnabled = $true; assignedLicenses = @(); userType = 'Guest' }
    )
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Find-NSPGraphMailbox' {
    BeforeEach {
        $users = $script:users
        Set-NSPGraphTransport -Name Fake -Responder { @{ value = $users } }.GetNewClosure()
    }
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'searches name, mail and UPN with eventual consistency and reads nothing else' {
        Find-NSPGraphMailbox -Search 'scan' | Out-Null
        $calls = @((Get-NSPGraphTransport).CallLog)
        $calls.Count | Should -Be 1
        $calls[0].Method | Should -Be 'GET'
        $calls[0].Headers.ConsistencyLevel | Should -Be 'eventual'
        $decoded = [uri]::UnescapeDataString($calls[0].Uri)
        $decoded | Should -BeLike 'v1.0/users?$search="displayName:scan" OR "mail:scan" OR "userPrincipalName:scan"*'
        $decoded | Should -Match '\$select=id,displayName,mail,userPrincipalName,accountEnabled,assignedLicenses,userType'
    }

    It 'leaves out accounts without mail and guests, and puts likely shared mailboxes first' {
        $found = @(Find-NSPGraphMailbox -Search 'scan')
        $found.Mail | Should -Be @('scanner@contoso.example', 'scan.op@contoso.example')
        $found[0].MailboxHint | Should -Be 'LikelyShared'
        $found[1].MailboxHint | Should -Be 'LicensedUser'
        $found[1].Licensed | Should -BeTrue
    }

    It 'resolves an exact address through mail, UPN or SMTP proxy, escaping quotes' {
        Find-NSPGraphMailbox -Address "o'neil@contoso.example" | Out-Null
        $decoded = [uri]::UnescapeDataString(@((Get-NSPGraphTransport).CallLog)[0].Uri)
        $decoded | Should -Match "mail eq 'o''neil@contoso.example' or userPrincipalName eq 'o''neil@contoso.example' or proxyAddresses/any\(p:p eq 'smtp:o''neil@contoso.example'\)"
    }

    It 'rejects a search term that would break the query' {
        { Find-NSPGraphMailbox -Search 'scan"x' } | Should -Throw
        @((Get-NSPGraphTransport).CallLog).Count | Should -Be 0
    }
}
