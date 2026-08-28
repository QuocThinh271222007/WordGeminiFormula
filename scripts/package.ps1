param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$project = Join-Path $root 'src\WordGeminiFormula.AddIn\WordGeminiFormula.AddIn.csproj'
$dist = Join-Path $root 'dist\WordGeminiFormula'

if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Path $dist -Force | Out-Null

msbuild $project /t:Restore,Build /p:Configuration=$Configuration /p:Platform='AnyCPU'
Copy-Item (Join-Path $root "src\WordGeminiFormula.AddIn\bin\$Configuration\net48\WordGeminiFormula.AddIn.dll") $dist
Copy-Item (Join-Path $root 'scripts\install.ps1') $dist
Copy-Item (Join-Path $root 'scripts\uninstall.ps1') $dist
Copy-Item (Join-Path $root 'README.md') $dist

Compress-Archive -Path "$dist\*" -DestinationPath (Join-Path $root 'dist\WordGeminiFormula.zip') -Force
Write-Host "Package: $(Join-Path $root 'dist\WordGeminiFormula.zip')"
