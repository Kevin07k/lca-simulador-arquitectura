# Shared list: import order of the LCA VBA modules (src/vba). Add new modules here.
$script:LcaModules = @(
    'LcaInicio', 'LcaCpu', 'LcaAlu', 'LcaBus', 'LcaControl',
    'LcaRam', 'LcaMemoryPath', 'LcaCache', 'LcaLayout',
    'LcaPanel', 'LcaDiagrama', 'LcaBusPanel', 'LcaCachePanel',
    'LcaCpuTests', 'LcaBusTests', 'LcaRamTests', 'LcaCacheTests'
)

function Get-RepoRoot { (Resolve-Path (Join-Path $PSScriptRoot '..')).Path }

# Opens Excel and returns it; fails with a clear message if VBA project access is off.
function Start-LcaExcel {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    return $excel
}

function Assert-VbaAccess($workbook) {
    $ok = $false
    try { $ok = $null -ne $workbook.VBProject -and $null -ne $workbook.VBProject.VBComponents } catch { $ok = $false }
    if (-not $ok) {
        throw ("Excel bloquea el acceso al proyecto VBA / Excel blocks VBA project access.`n" +
               "Activar manualmente / Enable manually: Excel > Archivo > Opciones > Centro de confianza > " +
               "Configuracion del Centro de confianza > Configuracion de macros > " +
               "'Confiar en el acceso al modelo de objetos de proyectos de VBA' / 'Trust access to the VBA project object model'.")
    }
}
