param(
    [Parameter(Mandatory = $false)]
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-Reg([string[]]$Arguments) {
    & "$env:SystemRoot\System32\reg.exe" @Arguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "reg.exe failed with exit code $LASTEXITCODE: $($Arguments -join ' ')"
    }
}

if ([string]::IsNullOrWhiteSpace($DllPath)) {
    $packagedDll = Join-Path $PSScriptRoot 'WordGeminiFormula.AddIn.dll'
    $sourceDll = Join-Path $PSScriptRoot '..\src\WordGeminiFormula.AddIn\bin\Release\net48\WordGeminiFormula.AddIn.dll'

    if (Test-Path $packagedDll) {
        $DllPath = $packagedDll
    }
    elseif (Test-Path $sourceDll) {
        $DllPath = $sourceDll
    }
    else {
        throw 'WordGeminiFormula.AddIn.dll was not found. Build Release first when using source, or extract the complete artifact ZIP and keep the DLL next to install.ps1.'
    }
}

$DllPath = [System.IO.Path]::GetFullPath($DllPath)
if (-not (Test-Path $DllPath)) {
    throw "DLL not found: $DllPath"
}

if (-not (Test-IsAdministrator)) {
    Write-Host 'Administrator permission is required for .NET COM registration. Opening UAC...' -ForegroundColor Yellow
    $quotedScript = '"' + $PSCommandPath.Replace('"', '\"') + '"'
    $quotedDll = '"' + $DllPath.Replace('"', '\"') + '"'
    $argLine = "-NoProfile -ExecutionPolicy Bypass -File $quotedScript -DllPath $quotedDll"
    $process = Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -Verb RunAs -ArgumentList $argLine -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Elevated installer failed with exit code $($process.ExitCode)."
    }
    exit 0
}

$clsid = '{7BA1B881-3DA4-4FBA-A25D-5F92141658EE}'
$progId = 'WordGeminiFormula.AddIn'
$officeAddinKey = "HKCU\Software\Microsoft\Office\Word\Addins\$progId"

# Clean legacy per-user COM registration created by V0.1.
$legacyKeys = @(
    "HKCU\Software\Classes\CLSID\$clsid",
    "HKCU\Software\Classes\$progId"
)
foreach ($key in $legacyKeys) {
    & "$env:SystemRoot\System32\reg.exe" delete $key /f /reg:64 2>$null | Out-Null
    & "$env:SystemRoot\System32\reg.exe" delete $key /f /reg:32 2>$null | Out-Null
}

# Register the managed COM class in both registry views on 64-bit Windows so
# either 32-bit or 64-bit Word can load the same AnyCPU .NET Framework assembly.
$regasmPaths = New-Object System.Collections.Generic.List[string]
if ([Environment]::Is64BitOperatingSystem) {
    $regasm64 = "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    if (Test-Path $regasm64) { $regasmPaths.Add($regasm64) }
}
$regasm32 = "$env:SystemRoot\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
if (Test-Path $regasm32) { $regasmPaths.Add($regasm32) }

if ($regasmPaths.Count -eq 0) {
    throw '.NET Framework RegAsm.exe was not found. Install/repair .NET Framework 4.8.'
}

foreach ($regasm in $regasmPaths) {
    & $regasm $DllPath /nologo /codebase
    if ($LASTEXITCODE -ne 0) {
        throw "RegAsm failed with exit code $LASTEXITCODE: $regasm"
    }
}

# Word discovers COM add-ins through this per-user key. Write both registry
# views because Office may be installed as either 32-bit or 64-bit.
$views = if ([Environment]::Is64BitOperatingSystem) { @('64', '32') } else { @('32') }
foreach ($view in $views) {
    Invoke-Reg @('add', $officeAddinKey, '/v', 'FriendlyName', '/t', 'REG_SZ', '/d', 'Word Gemini Formula', '/f', "/reg:$view")
    Invoke-Reg @('add', $officeAddinKey, '/v', 'Description', '/t', 'REG_SZ', '/d', 'Gemini OCR and native Word equation normalization', '/f', "/reg:$view")
    Invoke-Reg @('add', $officeAddinKey, '/v', 'LoadBehavior', '/t', 'REG_DWORD', '/d', '3', '/f', "/reg:$view")
    Invoke-Reg @('add', $officeAddinKey, '/v', 'CommandLineSafe', '/t', 'REG_DWORD', '/d', '0', '/f', "/reg:$view")
}

Write-Host 'Word Gemini Formula COM registration completed successfully.' -ForegroundColor Green
Write-Host "DLL: $DllPath"
Write-Host 'Close all Word windows and reopen Word. Ribbon tab: AI Formula.'
