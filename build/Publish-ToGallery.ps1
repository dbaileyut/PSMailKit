$ModuleRoot = (Get-Item (Join-Path -Path $PSScriptRoot '..')).FullName
Invoke-ScriptAnalyzer -Recurse -Severity Warning -Path $ModuleRoot