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
    Subject = "#$TestEmailNum $TestTime Send-MKMailMerge - Test Plain Parameters"
    MessageTemplate = Join-Path -Path $ModuleRootPath $TestSettings.TemplateTextFile
    Csv = Join-Path -Path $ModuleRootPath $TestSettings.Csv
    SMTPServer = $TestSettings.SMTPServer
}
Send-MKMailMerge @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMerge - Test TestAddress"
$TestParams['TestAddress'] = $TestSettings.To
Send-MKMailMerge @TestParams
$TestEmailNum++
$TestParams.Remove('TestAddress')

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMerge - Test HTML Template"
$TestParams['MessageTemplate'] = Join-Path -Path $ModuleRootPath $TestSettings.TemplateHtmlFile
Send-MKMailMerge @TestParams
$TestEmailNum++

$TestParams['Subject'] = "#$TestEmailNum $TestTime Send-MKMailMerge - Test Authentication"
$TestParams['Credential'] = $SMTPCredential
$TestParams['SmtpServer'] = $TestSettings.AuthenticatedSmtpServer
$TestParams['Port'] = $TestSettings.AuthenticatedSmtpPort
Send-MKMailMerge @TestParams
$TestEmailNum++