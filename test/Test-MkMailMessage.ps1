$ModuleRootPath = Split-Path -Parent $PSScriptRoot
Import-Module $ModuleRootPath\PSMailKit.psd1 -Force
$TestSettings = Get-Content $PSScriptRoot\TestSettings.json | ConvertFrom-Json
$TestTime = Get-Date -Format "s"
"Test Time: $TestTime"
$TestEmailNum = 1
if (-not $SMTPCredential) {
    $SMTPCredential = Get-Credential -Message "Enter your SMTP credentials"
}
$TestParams = @{
    From = $TestSettings.From
    To = $TestSettings.ToMultiple
    Cc = $TestSettings.Cc
    Bcc = $TestSettings.Bcc
    Subject = "#$TestEmailNum $TestTime Send-MKMailMessage - Test Plain Parameters"
    Body = (Get-Content (Join-Path -Path $ModuleRootPath $TestSettings.BodyTextFile) -Raw)
    SMTPServer = $TestSettings.SMTPServer
}
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Test Plain Parameters with Port"
$TestParams['Port'] = $TestSettings.SMTPPort
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Test Plain Parameters with Attachment"
$TestParams['Attachments'] = $TestSettings.Attachments | % { Join-Path -Path $ModuleRootPath $_}
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Test Authenticated Parameters"
$TestParams['SmtpServer'] = $TestSettings.AuthenticatedSmtpServer
$TestParams['Port'] = $TestSettings.AuthenticatedSmtpPort
$TestParams['Credential'] = $SMTPCredential
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Test HTML Parameters"
$TestParams['Body'] = (Get-Content (Join-Path -Path $ModuleRootPath $TestSettings.BodyHtmlFile) -Raw)
$TestParams['BodyAsHtml'] = $true
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Signed"
$TestParams['SMIME'] = 'Sign'
Send-MKMailMessage @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMessage - Signed"
$TestParams['To'] = $TestSettings.EncryptedTo
$TestParams.Remove('Cc')
$TestParams.Remove('Bcc')
$TestParams['SMIME'] = 'SignAndEncrypt'
Send-MKMailMessage @TestParams
$TestEmailNum++
