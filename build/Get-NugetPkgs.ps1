<#
.SYNOPSIS
    Downloads the required nuget packages for the module
.DESCRIPTION
    Must test after each nuget package update. This script will download the nuget packages and copy them to the resources folder,
    keeping the folder structure and only the required assemblies in the module's .psd1.

    Requires nuget.exe to be in the path or in the resources folder.
    #>

$Framework = 'net48'
$ModuleRoot = (Get-Item (Join-Path -Path $PSScriptRoot '..')).FullName
$ResourcesPath = Join-Path -Path $ModuleRoot 'resources'
$ModulePsdPath = Join-Path -Path $ModuleRoot "$(Split-Path -Leaf $ModuleRoot).psd1"

Push-Location $ResourcesPath
if (Test-Path '.\nuget.exe') {
    $NugetPath = '.\nuget.exe'
}
$NugetCmd = get-command 'nuget.exe' -ErrorAction Ignore

if ($NugetCmd) {
    $NugetPath = $NugetCmd.Path
}

Remove-Item .\nuget\* -Recurse
if (Test-Path '.\nugetmin') {
    Remove-Item .\nugetmin -Recurse
}

& $NugetPath install MailKit -NoCache -NonInteractive -ConfigFile .\nuget.config -Framework $Framework
# Powereshell 5.1 seems to require the older version of System.Runtime.CompilerServices.Unsafe
& $NugetPath install System.Runtime.CompilerServices.Unsafe -Version 4.5.3 -NoCache -NonInteractive -ConfigFile .\nuget.config -Framework $Framework

$content = Get-Content -Path $ModulePsdPath -Raw -ErrorAction Stop
$scriptBlock = [scriptblock]::Create( $content )
$scriptBlock.CheckRestrictedLanguage( $allowedCommands, [string[]]'PSScriptRoot', $true )
$ModuleHash = ( & $scriptBlock )

$DLLs = $ModuleHash.RequiredAssemblies | % {$_ -replace '\\resources\\', ''}
$DLLs | % {
    $Destination = "$($_ -replace 'nuget','nugetmin')"
    New-Item -Path ($Destination | Split-Path -Parent) -ItemType Directory -Force | Out-Null
    Copy-Item -Path $_ -Destination $Destination -Force
}

Remove-Item .\nuget -Recurse
Move-Item .\nugetmin .\nuget

Pop-Location