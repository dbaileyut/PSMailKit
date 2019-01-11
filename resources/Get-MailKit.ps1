Push-Location $PSScriptRoot
$NugetCmd = get-command 'nuget.exe'

if (-not $NugetCmd) {
    Write-Error "Nuget.exe not in current `$env:Path. Exiting"
    #return
}

& '..\..\..\downloads\nuget.exe' install MailKit

dir .\*\lib\* -Directory | ? {$_.Name -notmatch "^(net45$|portable-)"} | Remove-Item -Recurse

Pop-Location