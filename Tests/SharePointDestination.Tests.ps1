[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Fake transport responders take the transport''s fixed Method/Uri/Body signature.')]
param()

BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop
    $script:clientId = '33333333-3333-3333-3333-333333333333'
    function Get-Write { @((Get-NSPGraphTransport).CallLog | Where-Object Method -ne 'GET') }
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Find-NSPGraphSite' {
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'searches sites by name and returns IDs sorted by name' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ value = @(@{ id = 's2'; displayName = 'Scans B'; webUrl = 'https://contoso.sharepoint.com/sites/b' }, @{ id = 's1'; displayName = 'Scans A'; webUrl = 'https://contoso.sharepoint.com/sites/a' }) } }
        $found = @(Find-NSPGraphSite -Search 'Scans')
        $found.SiteId | Should -Be @('s1', 's2')
        [uri]::UnescapeDataString(@((Get-NSPGraphTransport).CallLog)[0].Uri) | Should -BeLike 'v1.0/sites?search=Scans&$select=*'
    }

    It 'resolves a link through the shares API with an unpadded base64url token' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ id = 'item-1'; name = 'Inbox'; webUrl = 'https://contoso.sharepoint.com/x'; folder = @{ childCount = 0 }; parentReference = @{ siteId = 'host,sc,web'; driveId = 'drive-1' } } }
        $url = 'https://contoso.sharepoint.com/sites/Scans/Shared Documents/Inbox?a=1'
        $hit = Find-NSPGraphSite -Url $url
        $hit.DriveId | Should -Be 'drive-1'; $hit.ItemId | Should -Be 'item-1'; $hit.SiteId | Should -Be 'host,sc,web'; $hit.IsFolder | Should -BeTrue
        $token = 'u!' + ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($url)).TrimEnd('=').Replace('/', '_').Replace('+', '-'))
        @((Get-NSPGraphTransport).CallLog)[0].Uri | Should -BeLike "v1.0/shares/$token/driveItem*"
        $token | Should -Not -Match '[=/+]'
    }

    It 'finds the site behind a library' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ id = 'root-1'; webUrl = 'https://contoso.sharepoint.com/lib'; parentReference = @{ siteId = 'host,sc,web' } } }
        (Find-NSPGraphSite -DriveId 'drive-1').SiteId | Should -Be 'host,sc,web'
        @((Get-NSPGraphTransport).CallLog)[0].Uri | Should -BeLike 'v1.0/drives/drive-1/root*'
    }

    It 'rejects a link that is not SharePoint' {
        { Find-NSPGraphSite -Url 'https://example.com/file' } | Should -Throw
    }
}

Describe 'Get-NSPGraphDriveFolder' {
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'lists document libraries only' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ value = @(@{ id = 'd1'; name = 'Documents'; driveType = 'documentLibrary' }, @{ id = 'd2'; name = 'Preservation'; driveType = 'other' }) } }
        $libraries = @(Get-NSPGraphDriveFolder -SiteId 'host,sc,web')
        $libraries.DriveId | Should -Be @('d1')
        $libraries[0].ItemId | Should -Be 'root'
    }

    It 'lists folders only under the given item' {
        Set-NSPGraphTransport -Name Fake -Responder { @{ value = @(@{ id = 'f1'; name = 'Inbox'; folder = @{ childCount = 2 } }, @{ id = 'x1'; name = 'note.pdf'; file = @{} }) } }
        $folders = @(Get-NSPGraphDriveFolder -DriveId 'd1' -ItemId 'parent-1')
        $folders.ItemId | Should -Be @('f1')
        $folders[0].DriveId | Should -Be 'd1'; $folders[0].ChildCount | Should -Be 2
        @((Get-NSPGraphTransport).CallLog)[0].Uri | Should -BeLike 'v1.0/drives/d1/items/parent-1/children*'
    }
}

Describe 'New-NSPGraphDriveFolder' {
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'creates a folder with fail-on-conflict and returns its IDs' {
        Set-NSPGraphTransport -Name Fake -Responder { param($Method, $Uri, $Body) @{ id = 'new-1'; name = $Body.name } }
        $folder = New-NSPGraphDriveFolder -DriveId 'd1' -Name 'ScanBridge Test'
        $folder.ItemId | Should -Be 'new-1'
        $call = @((Get-NSPGraphTransport).CallLog)[0]
        $call.Method | Should -Be 'POST'; $call.Uri | Should -Be 'v1.0/drives/d1/items/root/children'
        $call.Body.'@microsoft.graph.conflictBehavior' | Should -Be 'fail'
    }

    It 'creates nothing with -WhatIf and rejects illegal names' {
        Set-NSPGraphTransport -Name Fake -Responder { @{} }
        New-NSPGraphDriveFolder -DriveId 'd1' -Name 'Test' -WhatIf
        Get-Write | Should -BeNullOrEmpty
        { New-NSPGraphDriveFolder -DriveId 'd1' -Name 'a/b' } | Should -Throw
    }
}

Describe 'Grant-NSPGraphSitePermission' {
    AfterEach { Set-NSPGraphTransport -Name MgGraphRequest }

    It 'plans, then grants write to this app only' {
        $clientId = $script:clientId
        Set-NSPGraphTransport -Name Fake -Responder {
            param($Method, $Uri, $Body)
            if ($Method -eq 'GET') { return @{ value = @(@{ id = 'other'; roles = @('read'); grantedToIdentitiesV2 = @(@{ application = @{ id = 'someone-else' } }) }) } }
            @{ id = 'perm-new'; roles = $Body.roles }
        }.GetNewClosure()
        (Grant-NSPGraphSitePermission -SiteId 'host,sc,web' -ClientId $clientId -DisplayName 'NSP-ScanBridge-Ingest').Status | Should -Be 'PlanOnly'
        Get-Write | Should -BeNullOrEmpty
        $result = Grant-NSPGraphSitePermission -SiteId 'host,sc,web' -ClientId $clientId -DisplayName 'NSP-ScanBridge-Ingest' -Execute -Confirm:$false
        $result.Status | Should -Be 'Granted'; $result.PermissionId | Should -Be 'perm-new'
        $post = @(Get-Write)[0]
        $post.Method | Should -Be 'POST'
        $post.Body.roles | Should -Be @('write')
        $post.Body.grantedToIdentities[0].application.id | Should -Be $clientId
    }

    It 'upgrades a weaker grant for this app and leaves an equal or stronger one alone' {
        $clientId = $script:clientId
        $state = @{ Roles = @('read') }
        Set-NSPGraphTransport -Name Fake -Responder {
            param($Method, $Uri, $Body)
            if ($Method -eq 'GET') { return @{ value = @(@{ id = 'mine'; roles = $state.Roles; grantedToIdentities = @(@{ application = @{ id = $clientId } }) }) } }
            @{}
        }.GetNewClosure()
        (Grant-NSPGraphSitePermission -SiteId 's' -ClientId $clientId -DisplayName 'x' -Execute -Confirm:$false).Status | Should -Be 'Upgraded'
        $patch = @(Get-Write)[0]
        $patch.Method | Should -Be 'PATCH'; $patch.Uri | Should -Be 'v1.0/sites/s/permissions/mine'
        $state.Roles = @('write')
        Set-NSPGraphTransport -Name Fake -Responder {
            param($Method, $Uri, $Body)
            if ($Method -eq 'GET') { return @{ value = @(@{ id = 'mine'; roles = @('fullcontrol'); grantedToIdentitiesV2 = @(@{ application = @{ id = $clientId } }) }) } }
            throw 'no write expected'
        }.GetNewClosure()
        (Grant-NSPGraphSitePermission -SiteId 's' -ClientId $clientId -DisplayName 'x' -Execute -Confirm:$false).Status | Should -Be 'AlreadyGranted'
    }
}
