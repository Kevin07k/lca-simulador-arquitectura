<#
.SYNOPSIS
  Recreates the LCA Excel workbook from src/vba (Windows + desktop Excel).
.DESCRIPTION
  Copies a base workbook (default workbook/LCA.xlsm, or a new blank one), removes the old Lca* modules,
  imports every module from src/vba in the order of LcaModules.ps1, compiles-by-load, optionally builds the
  panels and saves the result. The base workbook is never modified unless -InPlace is passed.
.PARAMETER OutputPath   Destination .xlsm (default: workbook/LCA.build.xlsm).
.PARAMETER BasePath     Starting workbook (default: workbook/LCA.xlsm if present, else a blank workbook).
.PARAMETER BuildPanels  Also runs ConstruirPanelCpu, ConstruirPanelBuses and ConstruirPanelCaches.
.PARAMETER InPlace      Write over workbook/LCA.xlsm instead of the .build copy.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\Build-Workbook.ps1 -BuildPanels
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$BasePath,
    [switch]$BuildPanels,
    [switch]$InPlace
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'LcaModules.ps1')

$root   = Get-RepoRoot
$srcDir = Join-Path $root 'src\vba'
$main   = Join-Path $root 'workbook\LCA.xlsm'
if (-not $BasePath)   { $BasePath = $main }
if (-not $OutputPath) { $OutputPath = if ($InPlace) { $main } else { Join-Path $root 'workbook\LCA.build.xlsm' } }

foreach ($m in $script:LcaModules) {
    if (-not (Test-Path (Join-Path $srcDir "$m.bas"))) { throw "Missing module file: src\vba\$m.bas" }
}

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("lca-build-" + [guid]::NewGuid().ToString('N') + '.xlsm')
$excel = $null; $wb = $null
try {
    $excel = Start-LcaExcel
    if (Test-Path $BasePath) {
        Copy-Item $BasePath $tmp
        $wb = $excel.Workbooks.Open($tmp)
    } else {
        $wb = $excel.Workbooks.Add()
        $wb.SaveAs($tmp, 52)   # xlOpenXMLWorkbookMacroEnabled
    }
    Assert-VbaAccess $wb

    # Remove previous standard modules named Lca* so imports never create duplicates (LcaCpu1...).
    $components = $wb.VBProject.VBComponents
    for ($i = $components.Count; $i -ge 1; $i--) {
        $c = $components.Item($i)
        if ($c.Type -eq 1 -and $c.Name -like 'Lca*') { $components.Remove($c) }
    }
    foreach ($m in $script:LcaModules) {
        $null = $components.Import((Join-Path $srcDir "$m.bas"))
        Write-Host "Imported $m"
    }

    if ($BuildPanels) {
        foreach ($macro in 'ConstruirPanelCpu', 'ConstruirPanelBuses', 'ConstruirPanelCaches') {
            $null = $excel.Run("'$($wb.Name)'!$macro")
            Write-Host "Ran $macro"
        }
    }

    $wb.Save()
    $wb.Close($true); $wb = $null
    Copy-Item $tmp $OutputPath -Force
    Write-Host "OK -> $OutputPath"
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
