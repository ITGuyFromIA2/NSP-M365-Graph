BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop

    # In-memory mailbox: default Contacts folder (5 contacts, paged by 3), folder "Work" with one
    # child folder "Old". Batch sub-requests delete from it. ThrottleOnce ids answer 429 the first
    # time; Forbidden ids answer 403; Missing ids answer 404.
    function Get-FakeMailbox {
        param([int]$DefaultCount = 5, [string[]]$ThrottleOnce = @(), [string[]]$Forbidden = @(), [string[]]$Missing = @())
        $u = 'users/user%40contoso.example'
        $state = @{
            Contacts = @{
                "$u/contacts" = [System.Collections.Generic.List[string]]::new()
                "$u/contactFolders/f1/contacts" = [System.Collections.Generic.List[string]]::new()
                "$u/contactFolders/f1/childFolders/f2/contacts" = [System.Collections.Generic.List[string]]::new()
            }
            Throttle = @{}
            Forbidden = $Forbidden
            Missing = $Missing
            Batches = [System.Collections.Generic.List[object]]::new()
        }
        for ($i = 1; $i -le $DefaultCount; $i++) { $state.Contacts["$u/contacts"].Add("c$i") }
        $state.Contacts["$u/contactFolders/f1/contacts"].Add('w1')
        $state.Contacts["$u/contactFolders/f1/childFolders/f2/contacts"].AddRange([string[]]@('o1', 'o2'))
        foreach ($id in $ThrottleOnce) { $state.Throttle[$id] = $true }

        $responder = {
            param($Method, $Uri, $Body)
            $rel = $Uri -replace '^https://graph\.microsoft\.com/', '' -replace '^v1\.0/', ''
            $path = $rel -replace '\?.*$', ''
            $skip = if ($rel -match '[?&]skip=(\d+)') { [int]$Matches[1] } else { 0 }

            if ($Method -eq 'POST' -and $path -eq '$batch') {
                $state.Batches.Add($Body)
                $responses = foreach ($req in $Body.requests) {
                    $url = $req.url.TrimStart('/') -replace '/permanentDelete$', ''
                    $id = $url.Substring($url.LastIndexOf('/') + 1)
                    $base = $url.Substring(0, $url.LastIndexOf('/'))
                    if ($state.Throttle[$id]) { $state.Throttle[$id] = $false; @{ id = $req.id; status = 429; headers = @{ 'Retry-After' = '7' } }; continue }
                    if ($state.Forbidden -contains $id) { @{ id = $req.id; status = 403; body = @{ error = @{ message = 'Access is denied.' } } }; continue }
                    # Missing: deleted elsewhere after the inventory read it.
                    if ($state.Missing -contains $id -and $state.Contacts.ContainsKey($base)) { [void]$state.Contacts[$base].Remove($id) }
                    if ($state.Missing -contains $id -or -not $state.Contacts.ContainsKey($base) -or -not $state.Contacts[$base].Contains($id)) { @{ id = $req.id; status = 404 }; continue }
                    [void]$state.Contacts[$base].Remove($id)
                    @{ id = $req.id; status = 204 }
                }
                return @{ responses = @($responses) }
            }
            if ($path -eq 'users/user%40contoso.example/contactFolders') { return @{ value = @(@{ id = 'f1'; displayName = 'Work' }) } }
            if ($path -eq 'users/user%40contoso.example/contactFolders/f1/childFolders') { return @{ value = @(@{ id = 'f2'; displayName = 'Old' }) } }
            if ($path -eq 'users/user%40contoso.example/contactFolders/f1/childFolders/f2/childFolders') { return @{ value = @() } }
            if ($state.Contacts.ContainsKey($path)) {
                $ids = @($state.Contacts[$path])
                $page = @($ids | Select-Object -Skip $skip -First 3 | ForEach-Object { @{ id = $_; displayName = "Contact $_" } })
                $result = @{ value = $page }
                if ($skip + 3 -lt $ids.Count) { $result['@odata.nextLink'] = "https://graph.microsoft.com/v1.0/$path`?skip=$($skip + 3)" }
                return $result
            }
            throw "Fake mailbox has no route for $Method $Uri"
        }.GetNewClosure()
        @{ State = $state; Responder = $responder }
    }
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Remove-NSPGraphMailboxContact' {
    BeforeEach {
        Mock Start-Sleep -ModuleName NSP.M365.Graph { }
    }
    AfterEach {
        Set-NSPGraphTransport -Name MgGraphRequest
    }

    It 'inventories nested folders and pages, writes the backup, and deletes nothing with -WhatIf' {
        $mb = Get-FakeMailbox
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -WhatIf
        $r.Folders | Should -Be 3
        $r.Found | Should -Be 8
        $r.Deleted | Should -Be 0
        $r.Completed | Should -BeFalse
        $mb.State.Batches.Count | Should -Be 0
        $backup = Get-Content -LiteralPath $r.BackupPath -Raw | ConvertFrom-Json
        @($backup.Contacts).Count | Should -Be 8
        @($backup.Folders) | Should -Be @('Contacts', 'Contacts\Work', 'Contacts\Work\Old')
        @($backup.Contacts | Where-Object FolderPath -eq 'Contacts\Work\Old').Count | Should -Be 2
    }

    It 'deletes every contact in batches of at most 20 using folder-qualified URLs' {
        $mb = Get-FakeMailbox -DefaultCount 25
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -Confirm:$false
        $r.Found | Should -Be 28
        $r.Deleted | Should -Be 28
        $r.Remaining | Should -Be 0
        $r.Completed | Should -BeTrue
        $mb.State.Batches.Count | Should -Be 2
        @($mb.State.Batches[0].requests).Count | Should -Be 20
        $urls = @($mb.State.Batches | ForEach-Object { $_.requests } | ForEach-Object { $_.url })
        $urls | Should -Contain '/users/user%40contoso.example/contacts/c1'
        $urls | Should -Contain '/users/user%40contoso.example/contactFolders/f1/childFolders/f2/contacts/o2'
        @($mb.State.Batches | ForEach-Object { $_.requests } | Where-Object method -ne 'DELETE').Count | Should -Be 0
    }

    It 'keeps requests a JSON array for a single contact' {
        $mb = Get-FakeMailbox -DefaultCount 1
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $null = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -BatchSize 1 -Confirm:$false
        (ConvertTo-Json -InputObject $mb.State.Batches[0] -Depth 8 -Compress) | Should -Match '"requests":\['
    }

    It 'retries throttled items after Retry-After, counts 404 as gone and reports 403 as failed' {
        $mb = Get-FakeMailbox -ThrottleOnce 'c2', 'o1' -Forbidden 'c4' -Missing 'w1'
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -Confirm:$false
        $r.Deleted | Should -Be 6
        $r.AlreadyGone | Should -Be 1
        @($r.Failed).Count | Should -Be 1
        $r.Failed[0].Id | Should -Be 'c4'
        $r.Failed[0].Status | Should -Be 403
        $r.Failed[0].Message | Should -Be 'Access is denied.'
        $r.Remaining | Should -Be 1
        $r.Completed | Should -BeFalse
        Should -Invoke Start-Sleep -ModuleName NSP.M365.Graph -ParameterFilter { $Seconds -eq 7 } -Times 1 -Exactly
        @($mb.State.Batches[1].requests).Count | Should -Be 2
    }

    It 'uses permanentDelete without a request body when -Permanent is set' {
        $mb = Get-FakeMailbox -DefaultCount 2
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -Permanent -Confirm:$false
        $r.Mode | Should -Be 'PermanentDelete'
        $r.Completed | Should -BeTrue
        $reqs = @($mb.State.Batches[0].requests)
        @($reqs | Where-Object { $_.method -ne 'POST' -or $_.url -notmatch '/permanentDelete$' }).Count | Should -Be 0
        @($reqs | Where-Object { $_.Contains('body') }).Count | Should -Be 0
    }

    It 'returns without a batch when the mailbox has no contacts' {
        $mb = Get-FakeMailbox -DefaultCount 0
        foreach ($k in @($mb.State.Contacts.Keys)) { $mb.State.Contacts[$k].Clear() }
        Set-NSPGraphTransport -Name Fake -Responder $mb.Responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -Confirm:$false
        $r.Found | Should -Be 0
        $r.Completed | Should -BeTrue
        $mb.State.Batches.Count | Should -Be 0
        Test-Path -LiteralPath $r.BackupPath | Should -BeTrue
    }

    It 'retries a GET that Graph throttles' {
        $mb = Get-FakeMailbox -DefaultCount 1
        $flag = @{ Done = $false }
        $inner = $mb.Responder
        $responder = {
            param($Method, $Uri, $Body)
            if (-not $flag.Done -and $Uri -match 'contactFolders\?') {
                $flag.Done = $true
                $ex = [System.Exception]::new('Too many requests')
                $ex | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 429; Headers = [pscustomobject]@{ RetryAfter = $null } })
                throw $ex
            }
            & $inner $Method $Uri $Body
        }.GetNewClosure()
        Set-NSPGraphTransport -Name Fake -Responder $responder
        $r = Remove-NSPGraphMailboxContact -UserId 'user@contoso.example' -BackupDirectory (Join-Path $TestDrive ([guid]::NewGuid().ToString('n'))) -WhatIf
        $r.Found | Should -Be 4
        Should -Invoke Start-Sleep -ModuleName NSP.M365.Graph -ParameterFilter { $Seconds -eq 2 } -Times 1 -Exactly
    }
}
