$ModuleRoot = (Get-Item (Join-Path -Path $PSScriptRoot '..')).FullName
$WarningsAndErrors = Invoke-ScriptAnalyzer -Recurse -Severity Warning -Path $ModuleRoot
if ($WarningsAndErrors) {
    Write-Error -Message "Script Analyzer found $($WarningsAndErrors.Count) warnings and errors. See output for details."
    $WarningsAndErrors | Format-List -Force
    return
}

if (-not $NugetApiKey) {
    $Credential = Get-Credential -Message "Enter your NuGet API key" -UserName "NuGet API Key"
    $NugetApiKey = $Credential.GetNetworkCredential().Password
}
# Doh PSMailKit name is already taken on the gallery.  Need to rename the module. Asked the gallery owner to let me take it.
Publish-Module -Path $ModuleRoot -NuGetApiKey $NugetApiKey -Repository PSGallery -Verbose
