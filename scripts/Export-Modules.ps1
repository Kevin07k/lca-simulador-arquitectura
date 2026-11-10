<#
.SYNOPSIS
  Exports every Lca* standard module from a workbook back to src/vba (keeps source and workbook aligned).
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\Export-Modules.ps1 -WorkbookPath workbook\LCA.xlsm
#>
[CmdletBinding()]
param([string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'LcaModules.ps1')

$root = Get-RepoRoot
if (-not $WorkbookPath) { $WorkbookPath = Join-Path $root 'workbook\LCA.xlsm' }
$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$srcDir = Join-Path $root 'src\vba'

$excel = $null; $wb = $null
try {
    $excel = Start-LcaExcel
    $wb = $excel.Workbooks.Open($WorkbookPath, 0, $true)   # read-only
    Assert-VbaAccess $wb
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq 1 -and $c.Name -like 'Lca*') {
            $c.Export((Join-Path $srcDir "$($c.Name).bas"))
            Write-Host "Exported $($c.Name)"
        }
    }
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
Write-Host "Review with: git diff src/vba"
