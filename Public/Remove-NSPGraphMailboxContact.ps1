function Remove-NSPGraphMailboxContact {
    <#
    .SYNOPSIS
        Backs up, then deletes, every personal contact in one mailbox: the default Contacts folder
        and all of its subfolders.
    .DESCRIPTION
        Intended for clearing a mailbox before a contact re-import.

        1. Inventory: the default Contacts folder and every contact folder below it, recursively.
        2. Backup: every contact, as returned by Graph, is written to a JSON file and the file's
           count is checked before anything is deleted. This always happens, also with -WhatIf.
        3. Delete: in $batch requests of up to 20. Sub-requests answered 429 or 5xx are retried
           after the longest Retry-After in that round (minimum 5 s); 404 counts as already gone.
        4. Read-back: counts what is left, re-reading after 5, 15 and 30 s because Graph reads can
           lag writes.

        Default deletion moves contacts to Deleted Items. -Permanent uses permanentDelete, which
        places them in the Purges folder of Recoverable Items (not visible to mail clients; kept
        until the retention period or hold ends).

        Folders are kept. Hidden system folders that the contactFolders API does not list (for
        example Recipient Cache) are not touched. Contacts that sync to phones through Exchange
        ActiveSync disappear from those phones too.

        Needs a Graph session that can write the mailbox's contacts: app-only with the
        Contacts.ReadWrite application permission (any mailbox, or one scoped with Exchange RBAC
        for Applications), or delegated Contacts.ReadWrite for the signed-in user's own mailbox.
        This function does not connect.

        ConfirmImpact is High: PowerShell asks before deleting unless -Confirm:$false is passed.
    .OUTPUTS
        One object: UserId, Mode, Folders, Found, BackupPath, Deleted, AlreadyGone, Failed (the
        contacts that could not be deleted, with status and message), Remaining, Completed.
    .EXAMPLE
        Remove-NSPGraphMailboxContact -UserId user@contoso.example -WhatIf
        Inventory and backup only.
    .EXAMPLE
        Remove-NSPGraphMailboxContact -UserId user@contoso.example
        Backs up, asks, then moves every contact to Deleted Items.
    .EXAMPLE
        Remove-NSPGraphMailboxContact -UserId user@contoso.example -Permanent -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$UserId,
        [string]$BackupDirectory = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'NSP\ContactBackup'),
        [switch]$Permanent,
        [ValidateRange(1, 20)][int]$BatchSize = 20,
        [ValidateRange(0, 20)][int]$MaxRetry = 8
    )

    $user = [uri]::EscapeDataString($UserId)
    $mode = if ($Permanent) { 'PermanentDelete' } else { 'DeletedItems' }

    function Read-AllPages([string]$Uri) {
        $items = [System.Collections.Generic.List[object]]::new()
        $next = $Uri
        while ($next) {
            $page = Invoke-NSPGraphRetryGet -Uri $next -MaxRetry $MaxRetry
            foreach ($item in @($page['value'])) { if ($null -ne $item) { $items.Add($item) } }
            $next = $page['@odata.nextLink']
        }
        , $items.ToArray()
    }

    # --- 1. Inventory --------------------------------------------------------------------------
    $folders = [System.Collections.Generic.List[object]]::new()
    $folders.Add([pscustomobject]@{ Path = 'Contacts'; Base = "users/$user/contacts" })
    $queue = [System.Collections.Generic.Queue[object]]::new()
    foreach ($f in (Read-AllPages "users/$user/contactFolders?`$top=100&`$select=id,displayName")) {
        $queue.Enqueue([pscustomobject]@{ Path = "Contacts\$($f['displayName'])"; Folder = "users/$user/contactFolders/$($f['id'])" })
    }
    while ($queue.Count -gt 0) {
        $f = $queue.Dequeue()
        $folders.Add([pscustomobject]@{ Path = $f.Path; Base = "$($f.Folder)/contacts" })
        foreach ($c in (Read-AllPages "$($f.Folder)/childFolders?`$top=100&`$select=id,displayName")) {
            $queue.Enqueue([pscustomobject]@{ Path = "$($f.Path)\$($c['displayName'])"; Folder = "$($f.Folder)/childFolders/$($c['id'])" })
        }
    }

    $all = [System.Collections.Generic.List[object]]::new()
    foreach ($folder in $folders) {
        $contacts = Read-AllPages "$($folder.Base)?`$top=500"
        Write-Verbose ("{0}: {1} contact(s)" -f $folder.Path, $contacts.Count)
        foreach ($c in $contacts) { $all.Add([pscustomobject]@{ FolderPath = $folder.Path; Base = $folder.Base; Id = $c['id']; Contact = $c }) }
    }

    # --- 2. Backup -----------------------------------------------------------------------------
    $null = New-Item -ItemType Directory -Force -Path $BackupDirectory -WhatIf:$false -Confirm:$false
    $safeName = $UserId -replace '[^\w\.-]', '_'
    $backupPath = Join-Path $BackupDirectory ('{0}_{1:yyyyMMdd-HHmmss}.json' -f $safeName, [datetime]::Now)
    $backup = [ordered]@{
        UserId = $UserId
        CapturedUtc = [datetime]::UtcNow.ToString('o')
        Folders = @($folders | ForEach-Object { $_.Path })
        Contacts = @($all | ForEach-Object { [ordered]@{ FolderPath = $_.FolderPath; Contact = $_.Contact } })
    }
    [IO.File]::WriteAllText($backupPath, (ConvertTo-Json -InputObject $backup -Depth 16), [Text.UTF8Encoding]::new($false))
    $written = @((Get-Content -LiteralPath $backupPath -Raw | ConvertFrom-Json).Contacts).Count
    if ($all.Count -eq 0) { $written = 0 }
    if ($written -ne $all.Count) { throw "Backup check failed: $backupPath holds $written of $($all.Count) contacts. Nothing was deleted." }

    $result = [pscustomobject]@{
        UserId = $UserId
        Mode = $mode
        Folders = $folders.Count
        Found = $all.Count
        BackupPath = $backupPath
        Deleted = 0
        AlreadyGone = 0
        Failed = @()
        Remaining = $all.Count
        Completed = $false
    }
    if ($all.Count -eq 0) { $result.Completed = $true; return $result }

    $action = if ($Permanent) { "Permanently delete $($all.Count) contact(s) in $($folders.Count) folder(s)" } else { "Delete $($all.Count) contact(s) in $($folders.Count) folder(s) to Deleted Items" }
    if (-not $PSCmdlet.ShouldProcess($UserId, $action)) { return $result }

    # --- 3. Delete -----------------------------------------------------------------------------
    $pending = $all
    $failed = [System.Collections.Generic.List[object]]::new()
    $round = 0
    while ($pending.Count -gt 0) {
        $retry = [System.Collections.Generic.List[object]]::new()
        $wait = 0
        for ($o = 0; $o -lt $pending.Count; $o += $BatchSize) {
            $chunk = $pending.GetRange($o, [math]::Min($BatchSize, $pending.Count - $o))
            $requests = [System.Collections.Generic.List[object]]::new()
            for ($n = 0; $n -lt $chunk.Count; $n++) {
                $url = "/$($chunk[$n].Base)/$($chunk[$n].Id)"
                if ($Permanent) { $requests.Add([ordered]@{ id = "$n"; method = 'POST'; url = "$url/permanentDelete" }) }
                else { $requests.Add([ordered]@{ id = "$n"; method = 'DELETE'; url = $url }) }
            }
            try {
                # ToArray keeps requests a JSON array even for one element.
                $response = Invoke-NSPGraphRequest -Method POST -Uri '$batch' -Body @{ requests = $requests.ToArray() }
            } catch {
                $status = 0
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
                if ($status -ne 0 -and $status -ne 429 -and $status -lt 500) { throw }
                Write-Verbose "Batch request failed ($status); its $($chunk.Count) item(s) will be retried."
                $retry.AddRange($chunk); $wait = [math]::Max($wait, 10)
                continue
            }
            foreach ($r in @($response['responses'])) {
                $item = $chunk[[int]$r['id']]
                $status = [int]$r['status']
                if ($status -ge 200 -and $status -lt 300) { $result.Deleted++ }
                elseif ($status -eq 404) { $result.AlreadyGone++ }
                elseif ($status -eq 429 -or $status -ge 500) {
                    $retry.Add($item)
                    $after = 0
                    $headers = $r['headers']
                    if ($headers -and $headers['Retry-After']) { [void][int]::TryParse([string]$headers['Retry-After'], [ref]$after) }
                    $wait = [math]::Max($wait, [math]::Max($after, 5))
                } else {
                    $message = ''
                    $body = $r['body']
                    if ($body -and $body['error']) { $message = [string]$body['error']['message'] }
                    $failed.Add([pscustomobject]@{ FolderPath = $item.FolderPath; Id = $item.Id; Status = $status; Message = $message })
                }
            }
            $done = $result.Deleted + $result.AlreadyGone + $failed.Count
            Write-Progress -Activity "Deleting contacts for $UserId" -Status "$done of $($all.Count); $($retry.Count) to retry" -PercentComplete ([math]::Min(100, 100 * $done / $all.Count))
        }
        if ($retry.Count -eq 0) { break }
        $round++
        if ($round -gt $MaxRetry) {
            foreach ($item in $retry) { $failed.Add([pscustomobject]@{ FolderPath = $item.FolderPath; Id = $item.Id; Status = 429; Message = 'Retries exhausted' }) }
            break
        }
        Write-Verbose "$($retry.Count) item(s) throttled or transient; waiting $wait s (round $round of $MaxRetry)."
        Start-Sleep -Seconds $wait
        $pending = $retry
    }
    Write-Progress -Activity "Deleting contacts for $UserId" -Completed
    $result.Failed = $failed.ToArray()

    # --- 4. Read-back --------------------------------------------------------------------------
    $remaining = 0
    foreach ($delay in 0, 5, 15, 30) {
        if ($delay) { Start-Sleep -Seconds $delay }
        $remaining = 0
        foreach ($folder in $folders) { $remaining += (Read-AllPages "$($folder.Base)?`$select=id&`$top=500").Count }
        if ($remaining -le $failed.Count) { break }
    }
    $result.Remaining = $remaining
    $result.Completed = ($remaining -eq 0 -and $failed.Count -eq 0)
    $result
}
