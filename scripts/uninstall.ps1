$ErrorActionPreference = 'Stop'
$clsid = '{7BA1B881-3DA4-4FBA-A25D-5F92141658EE}'
$progId = 'WordGeminiFormula.AddIn'

$paths = @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid"
)

foreach ($path in $paths) {
    if (Test-Path $path) { Remove-Item -Path $path -Recurse -Force }
}

Write-Host 'Word Gemini Formula registration was removed. Close and reopen Word.' -ForegroundColor Green
