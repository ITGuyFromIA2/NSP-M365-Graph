function Invoke-NSPGraphRetryGet {
    <#
    .SYNOPSIS
        One Graph GET through Invoke-NSPGraphRequest, retried on 429 and transient 5xx.
    .DESCRIPTION
        Waits for Retry-After when the error carries one, otherwise 2, 4, 8 ... seconds (cap 60).
        Only GETs: a read is safe to repeat. Other errors are rethrown unchanged.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateRange(0, 20)][int]$MaxRetry = 8
    )

    $attempt = 0
    while ($true) {
        try {
            return Invoke-NSPGraphRequest -Method GET -Uri $Uri
        } catch {
            $attempt++
            $status = 0
            $wait = 0
            $response = $null
            try { $response = $_.Exception.Response } catch { $response = $null }
            if ($response) {
                try { $status = [int]$response.StatusCode } catch { $status = 0 }
                try {
                    $retryAfter = $response.Headers.RetryAfter
                    if ($retryAfter -and $retryAfter.Delta) { $wait = [int][math]::Ceiling($retryAfter.Delta.TotalSeconds) }
                } catch { $wait = 0 }
            }
            $transient = $status -eq 429 -or ($status -ge 500 -and $status -le 599)
            if (-not $transient -or $attempt -gt $MaxRetry) { throw }
            if ($wait -le 0) { $wait = [int][math]::Min(60, [math]::Pow(2, $attempt)) }
            Write-Verbose "Graph GET returned $status; retry $attempt of $MaxRetry in $wait s."
            Start-Sleep -Seconds $wait
        }
    }
}
