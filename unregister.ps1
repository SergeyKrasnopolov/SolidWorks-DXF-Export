#Requires -Version 5.1
#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'
if (Get-Process SLDWORKS -ErrorAction SilentlyContinue) { throw 'Close SOLIDWORKS before unregistration.' }
$assembly = Join-Path $PSScriptRoot 'bin\SolidWorksDxfExport.dll'
if (!(Test-Path $assembly)) { throw 'Keep the original bin folder until unregistration is complete.' }
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
& $regasm $assembly /unregister
if ($LASTEXITCODE -ne 0) { throw "Unregistration failed (exit code $LASTEXITCODE)." }
Write-Host 'Unregistered. You can now remove the project folder.'
