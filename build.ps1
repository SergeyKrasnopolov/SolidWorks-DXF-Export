#Requires -Version 5.1
param(
    [string]$SolidWorksDir = "$env:ProgramFiles\SOLIDWORKS Corp\SOLIDWORKS"
)
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path $compiler)) { throw 'Install .NET Framework 4.x (64-bit) first.' }
$names = @('SolidWorks.Interop.sldworks.dll', 'SolidWorks.Interop.swconst.dll', 'SolidWorks.Interop.swpublished.dll')
$references = @()
foreach ($name in $names) {
    $candidates = @(
        (Join-Path $SolidWorksDir "api\redist\$name"),
        (Join-Path $SolidWorksDir $name)
    )
    $found = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (!$found) { throw "Missing $name. Specify the SOLIDWORKS installation folder using -SolidWorksDir." }
    $references += $found
}
$output = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$assembly = Join-Path $output 'SolidWorksDxfExport.dll'
$arguments = @('/nologo', '/target:library', '/platform:x64', '/optimize+', '/codepage:65001',
    "/out:$assembly", '/reference:System.Windows.Forms.dll')
foreach ($reference in $references) { $arguments += "/reference:$reference" }
$arguments += (Join-Path $PSScriptRoot 'src\FrontViewAddin.cs')
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)." }
foreach ($reference in $references) { Copy-Item $reference $output -Force }
Write-Host "Built: $assembly"
Write-Host 'Next: close SOLIDWORKS and run register.ps1 in an administrator PowerShell.'
