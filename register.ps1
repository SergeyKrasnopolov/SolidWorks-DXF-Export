#Requires -Version 5.1
#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'
if (Get-Process SLDWORKS -ErrorAction SilentlyContinue) { throw 'Close SOLIDWORKS before registration.' }
$assembly = Join-Path $PSScriptRoot 'bin\SolidWorksDxfExport.dll'
if (!(Test-Path $assembly)) { throw 'Run build.ps1 first.' }
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
& $regasm $assembly /codebase
if ($LASTEXITCODE -ne 0) { throw "Registration failed (exit code $LASTEXITCODE)." }
Write-Host 'Registered. Open a model in SOLIDWORKS, then enable DXF Export - Front View under Tools > Add-Ins.'
