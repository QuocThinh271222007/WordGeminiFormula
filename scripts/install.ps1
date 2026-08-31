param(
    [Parameter(Mandatory = $false)]
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'

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

$clsid = '{7BA1B881-3DA4-4FBA-A25D-5F92141658EE}'
$progId = 'WordGeminiFormula.AddIn'
$className = 'WordGeminiFormula.AddIn.Connect'
$assemblyVersion = '0.1.0.0'
$assembly = "WordGeminiFormula.AddIn, Version=$assemblyVersion, Culture=neutral, PublicKeyToken=null"
$runtime = 'v4.0.30319'
$codeBase = ([System.Uri]$DllPath).AbsoluteUri
$managedCategory = '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

$clsidPath = "HKCU:\Software\Classes\CLSID\$clsid"
$inproc = Join-Path $clsidPath 'InprocServer32'
$versionedInproc = Join-Path $inproc $assemblyVersion
$progIdPath = "HKCU:\Software\Classes\$progId"
$addinPath = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"

# Remove stale class registration from earlier versions before recreating it.
if (Test-Path $clsidPath) { Remove-Item -Path $clsidPath -Recurse -Force }
if (Test-Path $progIdPath) { Remove-Item -Path $progIdPath -Recurse -Force }

New-Item -Path $inproc -Force | Out-Null
Set-Item -Path $clsidPath -Value $className -Force
Set-Item -Path $inproc -Value 'mscoree.dll' -Force
New-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Class' -Value $className -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Assembly' -Value $assembly -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value $runtime -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'CodeBase' -Value $codeBase -PropertyType String -Force | Out-Null

# RegAsm-style version-specific registration is required by the CLR COM loader.
New-Item -Path $versionedInproc -Force | Out-Null
New-ItemProperty -Path $versionedInproc -Name 'Class' -Value $className -PropertyType String -Force | Out-Null
New-ItemProperty -Path $versionedInproc -Name 'Assembly' -Value $assembly -PropertyType String -Force | Out-Null
New-ItemProperty -Path $versionedInproc -Name 'RuntimeVersion' -Value $runtime -PropertyType String -Force | Out-Null
New-ItemProperty -Path $versionedInproc -Name 'CodeBase' -Value $codeBase -PropertyType String -Force | Out-Null

New-Item -Path "$clsidPath\ProgId" -Force | Out-Null
Set-Item -Path "$clsidPath\ProgId" -Value $progId -Force
New-Item -Path "$clsidPath\Implemented Categories\$managedCategory" -Force | Out-Null

New-Item -Path "$progIdPath\CLSID" -Force | Out-Null
Set-Item -Path $progIdPath -Value $className -Force
Set-Item -Path "$progIdPath\CLSID" -Value $clsid -Force

New-Item -Path $addinPath -Force | Out-Null
New-ItemProperty -Path $addinPath -Name 'FriendlyName' -Value 'Word Gemini Formula' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinPath -Name 'Description' -Value 'Gemini OCR and native Word equation normalization' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinPath -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinPath -Name 'CommandLineSafe' -Value 0 -PropertyType DWord -Force | Out-Null

Write-Host 'Word Gemini Formula was registered for the current Windows user.' -ForegroundColor Green
Write-Host "DLL: $DllPath"
Write-Host 'Close all Word windows and reopen Word. Ribbon tab: AI Formula.'
