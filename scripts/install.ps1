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

function Get-RegistryViews {
    if ([Environment]::Is64BitOperatingSystem) {
        return @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32
        )
    }
    return @([Microsoft.Win32.RegistryView]::Registry32)
}

function Remove-HkcuSubKey([string]$SubKey) {
    foreach ($view in Get-RegistryViews) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view
        )
        try {
            $baseKey.DeleteSubKeyTree($SubKey, $false)
        }
        finally {
            $baseKey.Dispose()
        }
    }
}

function Set-WordAddinRegistry([string]$SubKey) {
    foreach ($view in Get-RegistryViews) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view
        )
        try {
            $key = $baseKey.CreateSubKey($SubKey)
            try {
                $key.SetValue('FriendlyName', 'Word Gemini Formula', [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('Description', 'Gemini OCR and native Word equation normalization', [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
                $key.SetValue('CommandLineSafe', 0, [Microsoft.Win32.RegistryValueKind]::DWord)
            }
            finally {
                $key.Dispose()
            }
        }
        finally {
            $baseKey.Dispose()
        }
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
$officeAddinSubKey = "Software\Microsoft\Office\Word\Addins\$progId"

# Clean legacy per-user COM registration created by V0.1. Missing keys are OK.
Remove-HkcuSubKey "Software\Classes\CLSID\$clsid"
Remove-HkcuSubKey "Software\Classes\$progId"

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
        throw "RegAsm failed with exit code ${LASTEXITCODE}: $regasm"
    }
}

# Word discovery remains per-user; write both registry views for 32/64-bit Office.
Set-WordAddinRegistry $officeAddinSubKey

Write-Host 'Word Gemini Formula COM registration completed successfully.' -ForegroundColor Green
Write-Host "DLL: $DllPath"
Write-Host 'Close all Word windows and reopen Word. Ribbon tab: AI Formula.'
