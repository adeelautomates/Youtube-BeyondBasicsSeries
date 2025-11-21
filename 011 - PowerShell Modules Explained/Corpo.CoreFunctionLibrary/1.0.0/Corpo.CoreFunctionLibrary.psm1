$functionsPath = Join-Path $PSScriptRoot 'Functions'
$files = Get-ChildItem $functionsPath -Filter '*.ps1'
$files | ForEach-Object { . $_.FullName }
Export-ModuleMember -Function $Files.BaseName