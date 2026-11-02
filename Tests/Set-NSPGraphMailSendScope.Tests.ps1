BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop

    # Exchange Online cmdlets do not exist offline. Stub the parameters this function uses so
    # Pester can mock them and filter on their arguments.
    $script:exchangeStubs = [ordered]@{
        'Get-Recipient' = 'Identity'
        'Get-ServicePrincipal' = 'Identity'
        'New-ServicePrincipal' = 'AppId', 'ObjectId', 'DisplayName'
        'Get-ManagementScope' = 'Identity'
        'New-ManagementScope' = 'Name', 'RecipientRestrictionFilter'
        'Set-ManagementScope' = 'Identity', 'RecipientRestrictionFilter'
        'Get-ManagementRoleAssignment' = 'RoleAssignee', 'Role'
        'New-ManagementRoleAssignment' = 'App', 'Role', 'CustomResourceScope'
        'Test-ServicePrincipalAuthorization' = 'Identity', 'Resource'
    }
    foreach ($name in $script:exchangeStubs.Keys) {
        $parameters = (@($script:exchangeStubs[$name]) | ForEach-Object { "`$$_" }) -join ', '
        Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create("[CmdletBinding()] param($parameters) throw 'Exchange stub called without a mock.'"))
    }

    $script:clientId = '11111111-1111-1111-1111-111111111111'
    $script:spId = '22222222-2222-2222-2222-222222222222'
    $script:scopeName = "NSP Mail.Send $script:clientId"

    function Use-FakeGraph {
        param([switch]$EntraMailSend)
        $grants = if ($EntraMailSend) { @(@{ resourceId = 'graph-sp'; appRoleId = 'role-mail-send' }) } else { @(@{ resourceId = 'graph-sp'; appRoleId = 'role-other' }) }
        $clientId = $script:clientId
        $spId = $script:spId
        Set-NSPGraphTransport -Name Fake -Responder {
            param($Method, $Uri)
            $path = ([uri]::UnescapeDataString($Uri) -replace '^v1\.0/', '') -replace '\?.*$', ''
            if ($Method -ne 'GET') { throw "Unexpected Graph write: $Method $path" }
            switch -Regex ($path) {
                '^servicePrincipals$' { return @{ value = @(@{ id = 'graph-sp'; appRoles = @(@{ value = 'Mail.Send'; id = 'role-mail-send' }, @{ value = 'User.Read.All'; id = 'role-other' }) }) } }
                '/appRoleAssignments$' { return @{ value = $grants } }
                '^servicePrincipals/' { return @{ id = $spId; appId = $clientId; displayName = 'NSP-ScanBridge-Alerts' } }
            }
            throw "Unexpected Graph read: $path"
        }.GetNewClosure()
    }
    function Invoke-Scope {
        param([switch]$Execute)
        $scopeArgs = @{ ClientId = $script:clientId; ServicePrincipalId = $script:spId; Mailbox = 'scanner@contoso.example'; Confirm = $false }
        if ($Execute) { $scopeArgs.Execute = $true }
        Set-NSPGraphMailSendScope @scopeArgs
    }
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
    foreach ($name in $script:exchangeStubs.Keys) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
}

Describe 'Set-NSPGraphMailSendScope' {
    BeforeEach {
        Use-FakeGraph
        Mock Get-Recipient { [pscustomobject]@{ RecipientTypeDetails = 'SharedMailbox'; PrimarySmtpAddress = 'scanner@contoso.example' } } -ModuleName NSP.M365.Graph
        Mock Get-ServicePrincipal { throw 'not found' } -ModuleName NSP.M365.Graph
        Mock Get-ManagementScope { throw 'not found' } -ModuleName NSP.M365.Graph
        Mock Get-ManagementRoleAssignment { @() } -ModuleName NSP.M365.Graph
        Mock New-ServicePrincipal { } -ModuleName NSP.M365.Graph
        Mock New-ManagementScope { } -ModuleName NSP.M365.Graph
        Mock Set-ManagementScope { } -ModuleName NSP.M365.Graph
        Mock New-ManagementRoleAssignment { } -ModuleName NSP.M365.Graph
        Mock Test-ServicePrincipalAuthorization { [pscustomobject]@{ RoleName = 'Application Mail.Send'; InScope = $true } } -ModuleName NSP.M365.Graph
    }
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'refuses while the app holds Graph Mail.Send in Entra, before touching Exchange' {
        Use-FakeGraph -EntraMailSend
        $result = Invoke-Scope -Execute
        $result.Status | Should -Be 'UnscopedEntraGrant'
        Should -Invoke Get-Recipient -Times 0 -ModuleName NSP.M365.Graph
        Should -Invoke New-ManagementRoleAssignment -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'plans the pointer, scope and assignment without changing anything' {
        $result = Invoke-Scope
        $result.Status | Should -Be 'PlanOnly'
        $result.Actions | Should -Be @('CreateExchangeServicePrincipal', 'CreateScope', 'CreateRoleAssignment')
        Should -Invoke New-ServicePrincipal -Times 0 -ModuleName NSP.M365.Graph
        Should -Invoke New-ManagementScope -Times 0 -ModuleName NSP.M365.Graph
        Should -Invoke New-ManagementRoleAssignment -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'scopes the app to the primary SMTP address only and verifies it' {
        Mock Get-Recipient { [pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; PrimarySmtpAddress = 'front.scanner@contoso.example' } } -ModuleName NSP.M365.Graph
        $result = Invoke-Scope -Execute
        $result.Status | Should -Be 'Scoped'
        $result.Verified | Should -BeTrue
        $result.PrimarySmtpAddress | Should -Be 'front.scanner@contoso.example'
        Should -Invoke New-ServicePrincipal -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter { $AppId -eq '11111111-1111-1111-1111-111111111111' -and $ObjectId -eq '22222222-2222-2222-2222-222222222222' }
        Should -Invoke New-ManagementScope -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
            $Name -eq 'NSP Mail.Send 11111111-1111-1111-1111-111111111111' -and $RecipientRestrictionFilter -eq "PrimarySmtpAddress -eq 'front.scanner@contoso.example'"
        }
        Should -Invoke New-ManagementRoleAssignment -Times 1 -ModuleName NSP.M365.Graph -ParameterFilter {
            $App -eq '11111111-1111-1111-1111-111111111111' -and $Role -eq 'Application Mail.Send' -and $CustomResourceScope -eq 'NSP Mail.Send 11111111-1111-1111-1111-111111111111'
        }
    }

    It 'updates only the scope filter when the mailbox changed' {
        Mock Get-ServicePrincipal { [pscustomobject]@{ Identity = 'sp-identity' } } -ModuleName NSP.M365.Graph
        Mock Get-ManagementScope { [pscustomobject]@{ RecipientFilter = "PrimarySmtpAddress -eq 'old@contoso.example'" } } -ModuleName NSP.M365.Graph
        Mock Get-ManagementRoleAssignment { [pscustomobject]@{ CustomResourceScope = $script:scopeName } } -ModuleName NSP.M365.Graph
        $result = Invoke-Scope -Execute
        $result.Actions | Should -Be @('UpdateScope')
        Should -Invoke Set-ManagementScope -Times 1 -ModuleName NSP.M365.Graph
        Should -Invoke New-ManagementRoleAssignment -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'reports AlreadyScoped and still verifies when nothing is missing' {
        Mock Get-ServicePrincipal { [pscustomobject]@{ Identity = 'sp-identity' } } -ModuleName NSP.M365.Graph
        Mock Get-ManagementScope { [pscustomobject]@{ RecipientFilter = "PrimarySmtpAddress -eq 'scanner@contoso.example'" } } -ModuleName NSP.M365.Graph
        Mock Get-ManagementRoleAssignment { [pscustomobject]@{ CustomResourceScope = $script:scopeName } } -ModuleName NSP.M365.Graph
        $result = Invoke-Scope
        $result.Status | Should -Be 'AlreadyScoped'
        $result.Verified | Should -BeTrue
        Should -Invoke Test-ServicePrincipalAuthorization -Times 1 -ModuleName NSP.M365.Graph
    }

    It 'leaves an existing assignment on another or no scope alone' {
        Mock Get-ServicePrincipal { [pscustomobject]@{ Identity = 'sp-identity' } } -ModuleName NSP.M365.Graph
        Mock Get-ManagementRoleAssignment { [pscustomobject]@{ CustomResourceScope = $null } } -ModuleName NSP.M365.Graph
        $result = Invoke-Scope -Execute
        $result.Status | Should -Be 'ConflictingAssignment'
        $result.Message | Should -Match 'unscoped'
        Should -Invoke New-ManagementScope -Times 0 -ModuleName NSP.M365.Graph
        Should -Invoke New-ManagementRoleAssignment -Times 0 -ModuleName NSP.M365.Graph
    }

    It 'reports an unverified scope instead of claiming success' {
        Mock Test-ServicePrincipalAuthorization { [pscustomobject]@{ RoleName = 'Application Mail.Send'; InScope = $false } } -ModuleName NSP.M365.Graph
        (Invoke-Scope -Execute).Verified | Should -BeFalse
    }

    It 'refuses a recipient that is not a user or shared mailbox' {
        Mock Get-Recipient { [pscustomobject]@{ RecipientTypeDetails = 'MailContact'; PrimarySmtpAddress = 'scanner@contoso.example' } } -ModuleName NSP.M365.Graph
        { Invoke-Scope -Execute } | Should -Throw '*not a user or shared mailbox*'
        Should -Invoke New-ManagementScope -Times 0 -ModuleName NSP.M365.Graph
    }
}
