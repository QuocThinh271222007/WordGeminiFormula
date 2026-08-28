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

Write-Host 'Đã gỡ đăng ký Word Gemini Formula. Đóng và mở lại Word.' -ForegroundColor Green
