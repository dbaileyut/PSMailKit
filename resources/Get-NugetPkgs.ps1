Push-Location $PSScriptRoot
$NugetCmd = get-command 'nuget.exe'

if (-not $NugetCmd) {
    Write-Error "Nuget.exe not in current `$env:Path. Exiting"
    #return
}

Remove-Item .\nuget\* -Recurse

nuget.exe install MailKit -NoCache -NonInteractive

Pop-Location