<#
.SYNOPSIS
  Runs every LCA engine regression (P04..P10) inside desktop Excel without dialogs.
.DESCRIPTION
  Opens a workbook (default workbook/LCA.build.xlsm, else workbook/LCA.xlsm) read-only in a hidden Excel,
  calls LcaVerificarTodoTexto (src/vba/LcaInicio.bas) and prints one PASS/FAIL line per suite.
  Exit code 0 only when every suite passes. The workbook is closed without saving.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\Build-Workbook.ps1 -BuildPanels
  powershell -ExecutionPolicy Bypass -File scripts\Test-Workbook.ps1
#>
[CmdletBinding()]
param([string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'LcaModules.ps1')

$root = Get-RepoRoot
if (-not $WorkbookPath) {
    $WorkbookPath = Join-Path $root 'workbook\LCA.build.xlsm'
    if (-not (Test-Path $WorkbookPath)) { $WorkbookPath = Join-Path $root 'workbook\LCA.xlsm' }
}
$WorkbookPath = (Resolve-Path $WorkbookPath).Path

# Work on a temporary copy so the tests (which rebuild report sheets) never touch the original.
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("lca-test-" + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item $WorkbookPath $tmp
$excel = $null; $wb = $null; $result = $null
try {
    $excel = Start-LcaExcel
    $wb = $excel.Workbooks.Open($tmp)
    $result = [string]$excel.Run("'$($wb.Name)'!LcaVerificarTodoTexto")
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
Write-Host "Workbook: $WorkbookPath"
Write-Host $result
if ($result -notlike 'ALL PASS*') { exit 1 }
