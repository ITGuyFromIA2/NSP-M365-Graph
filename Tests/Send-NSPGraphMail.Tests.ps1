BeforeAll {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'NSP.M365.Graph.psd1') -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module NSP.M365.Graph -Force -ErrorAction SilentlyContinue
}

Describe 'Send-NSPGraphMail' {
    BeforeEach {
        Set-NSPGraphTransport -Name Fake -Responder { @{} }
    }
    AfterEach {
        Set-NSPGraphTransport -Name MgGraphRequest
    }

    It 'posts one sendMail request as the From mailbox with recipients and body' {
        Send-NSPGraphMail -From 'reports@contoso.example' -To 'a@contoso.example', 'b@contoso.example' -Cc 'c@contoso.example' -Subject 'Nightly' -Body "Line 1`r`nLine 2"
        $calls = @((Get-NSPGraphTransport).CallLog)
        $calls.Count | Should -Be 1
        $calls[0].Method | Should -Be 'POST'
        $calls[0].Uri | Should -Be 'v1.0/users/reports%40contoso.example/sendMail'
        $m = $calls[0].Body.message
        @($m.toRecipients | ForEach-Object { $_.emailAddress.address }) | Should -Be @('a@contoso.example', 'b@contoso.example')
        $m.ccRecipients[0].emailAddress.address | Should -Be 'c@contoso.example'
        $m.ContainsKey('bccRecipients') | Should -BeFalse
        $m.body.contentType | Should -Be 'Text'
        $m.body.content | Should -Be "Line 1`r`nLine 2"
        $calls[0].Body.saveToSentItems | Should -BeFalse
        $m.ContainsKey('attachments') | Should -BeFalse
    }

    It 'attaches files as base64 fileAttachments' {
        $file = Join-Path $TestDrive 'run.log'
        [IO.File]::WriteAllText($file, 'log line')
        Send-NSPGraphMail -From 'r@contoso.example' -To 'o@contoso.example' -Bcc 'audit@contoso.example' -Subject 'Logs' -Body '' -Attachment $file -SaveToSentItems
        $call = @((Get-NSPGraphTransport).CallLog)[0]
        $a = @($call.Body.message.attachments)
        $a.Count | Should -Be 1
        $a[0].name | Should -Be 'run.log'
        $a[0].'@odata.type' | Should -Be '#microsoft.graph.fileAttachment'
        [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($a[0].contentBytes)) | Should -Be 'log line'
        $call.Body.saveToSentItems | Should -BeTrue
        $call.Body.message.bccRecipients[0].emailAddress.address | Should -Be 'audit@contoso.example'
        (ConvertTo-Json -InputObject $call.Body -Depth 8) | Should -Match '"toRecipients":\s*\[' -Because 'a single recipient must still serialize as an array'
    }

    It 'refuses attachments over the size limit before sending anything' {
        $file = Join-Path $TestDrive 'big.log'
        [IO.File]::WriteAllText($file, ('x' * 2000))
        { Send-NSPGraphMail -From 'r@contoso.example' -To 'o@contoso.example' -Subject 'S' -Body 'B' -Attachment $file -MaxAttachmentBytes 1000 } | Should -Throw '*over the 1000-byte limit*'
        @((Get-NSPGraphTransport).CallLog).Count | Should -Be 0
    }

    It 'sends nothing with -WhatIf' {
        Send-NSPGraphMail -From 'r@contoso.example' -To 'o@contoso.example' -Subject 'S' -Body 'B' -WhatIf
        @((Get-NSPGraphTransport).CallLog).Count | Should -Be 0
    }
}
