@{
    RootModule = 'NSP.M365.Graph.psm1'
    ModuleVersion = '0.1.1'
    GUID = '5608504a-bd61-4a57-9c87-d6af884a2aed'
    Author = 'Network Systems Plus'
    CompanyName = 'Network Systems Plus'
    Copyright = '(c) 2026 Network Systems Plus. All rights reserved.'
    Description = 'Shared Microsoft Graph connection, paging, and swappable request transport for NSP M365 tools.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Connect-NSPGraph'
        'Connect-NSPGraphAppOnly'
        'Get-NSPGraphTransport'
        'Invoke-NSPGraphCollection'
        'Invoke-NSPGraphRequest'
        'Register-NSPGraphAppOnlyRegistration'
        'Register-NSPGraphAppRegistration'
        'Send-NSPGraphMail'
        'Set-NSPGraphTransport'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('MicrosoftGraph', 'M365', 'Entra', 'NSP')
            ProjectUri = 'https://github.com/ITGuyFromIA2/NSP-Graph'
            LicenseUri = 'https://github.com/ITGuyFromIA2/NSP-Graph/blob/main/LICENSE'
            ReleaseNotes = '0.1.1: Send-NSPGraphMail - one sendMail request as a mailbox, with file attachments (size checked before sending).'
        }
    }
}
