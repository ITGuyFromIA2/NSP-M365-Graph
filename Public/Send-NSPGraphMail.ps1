function Send-NSPGraphMail {
    <#
    .SYNOPSIS
        Sends one message as a mailbox through Graph sendMail, with optional file attachments.
    .DESCRIPTION
        Uses the active connection (Connect-NSPGraph / Connect-NSPGraphAppOnly) and transport. An
        app-only connection needs the Mail.Send application permission, which reaches every mailbox
        in the tenant unless Exchange RBAC for Applications scopes it; a delegated one sends as the
        signed-in user or a mailbox they can send as.

        Attachments are files, sent inline (fileAttachment). Graph's sendMail request is limited to
        about 4 MB, so the attachments' total size is checked first (-MaxAttachmentBytes, default
        3 MB) and the call throws before sending anything when it is over. Throws when Graph
        rejects the message; returns nothing on success.
    .EXAMPLE
        Send-NSPGraphMail -From 'reports@contoso.example' -To 'ops@contoso.example' -Subject 'Nightly run' -Body 'All files sent.'
    .EXAMPLE
        Send-NSPGraphMail -From 'reports@contoso.example' -To 'ops@contoso.example' -Cc 'lead@contoso.example' -Subject 'Logs' -Body $text -Attachment (Get-ChildItem .\Logs -File).FullName
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string[]]$To,
        [string[]]$Cc,
        [string[]]$Bcc,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Body,
        [ValidateSet('Text', 'HTML')][string]$BodyType = 'Text',
        [string[]]$Attachment,
        [switch]$SaveToSentItems,
        [ValidateRange(1, 4000000)][int]$MaxAttachmentBytes = 3000000
    )

    # @() at each call site too: a one-address list would otherwise unroll to a bare hashtable and
    # serialize as an object, not the array Graph requires.
    $toRecipient = { param([string[]]$Address) @($Address | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } }) }
    $files = @($Attachment | Where-Object { $_ } | ForEach-Object { Get-Item -LiteralPath $_ -ErrorAction Stop })
    $total = ($files | Measure-Object -Property Length -Sum).Sum
    if ($total -gt $MaxAttachmentBytes) {
        throw "Attachments total $total bytes, over the $MaxAttachmentBytes-byte limit for one sendMail request. Send fewer or smaller files."
    }

    $message = @{
        subject = $Subject
        body = @{ contentType = $BodyType; content = $Body }
        toRecipients = @(& $toRecipient $To)
    }
    if ($Cc) { $message.ccRecipients = @(& $toRecipient $Cc) }
    if ($Bcc) { $message.bccRecipients = @(& $toRecipient $Bcc) }
    if ($files.Count) {
        $message.attachments = @($files | ForEach-Object {
                @{ '@odata.type' = '#microsoft.graph.fileAttachment'; name = $_.Name; contentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($_.FullName)) }
            })
    }

    if (-not $PSCmdlet.ShouldProcess(($To -join ', '), "Send '$Subject' from $From with $($files.Count) attachment(s)")) { return }
    Invoke-NSPGraphRequest -Method POST -Uri "users/$([uri]::EscapeDataString($From))/sendMail" -Body @{ message = $message; saveToSentItems = [bool]$SaveToSentItems } | Out-Null
}
