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
$excel = $null; $wb = $null; $result = $null; $driftReport = ""
try {
    $excel = Start-LcaExcel
    $wb = $excel.Workbooks.Open($tmp)
    $result = [string]$excel.Run("'$($wb.Name)'!LcaVerificarTodoTexto")
    # programs/arith/*.asm are documentation copies of LcaArith.bas: they must not drift from the workbook source.
    $count = [int]$excel.Run("'$($wb.Name)'!ArithCountValue")
    for ($i = 0; $i -lt $count; $i++) {
        $name = [string]$excel.Run("'$($wb.Name)'!ArithFile", $i)
        $source = ([string]$excel.Run("'$($wb.Name)'!ArithAsmText", $i)).Trim()
        $path = Join-Path $root (Join-Path 'programs\arith' $name)
        if (-not (Test-Path $path)) { $driftReport += "MISSING programs\arith\$name`n"; continue }
        $file = ((Get-Content $path -Raw) -replace "`r`n", "`n").Trim()
        if ($file -ne $source) { $driftReport += "DRIFT   programs\arith\$name differs from LcaArith.bas`n" }
    }
    # programs/echo/eco.asm is the documentation copy of LcaEcho.bas (P16).
    $name = [string]$excel.Run("'$($wb.Name)'!EchoFile")
    $source = ([string]$excel.Run("'$($wb.Name)'!EchoAsmText")).Trim()
    $path = Join-Path $root (Join-Path 'programs\echo' $name)
    if (-not (Test-Path $path)) { $driftReport += "MISSING programs\echo\$name`n" }
    else {
        $file = ((Get-Content $path -Raw) -replace "`r`n", "`n").Trim()
        if ($file -ne $source) { $driftReport += "DRIFT   programs\echo\$name differs from LcaEcho.bas`n" }
    }
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
Write-Host "Workbook: $WorkbookPath"
Write-Host $result
if ($driftReport.Length -gt 0) { Write-Host $driftReport; Write-Host "FAIL  programs\ differ from the workbook source" }
else { Write-Host "PASS  programs\arith\*.asm match LcaArith.bas; programs\echo\eco.asm matches LcaEcho.bas" }
if ($result -notlike 'ALL PASS*' -or $driftReport.Length -gt 0) { exit 1 }
