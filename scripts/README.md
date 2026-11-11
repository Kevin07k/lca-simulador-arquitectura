# Scripts de automatización / Automation scripts

Requisitos / Requirements: Windows, Excel de escritorio, PowerShell. Una sola vez, en Excel:
*Archivo > Opciones > Centro de confianza > Configuración > Configuración de macros > "Confiar en el acceso al modelo de objetos de proyectos de VBA"*. Los scripts no cambian esta opción.

| Script | Qué hace / What it does |
| --- | --- |
| `Build-Workbook.ps1` | Recrea el libro desde `src/vba`: copia `workbook/LCA.xlsm`, quita los módulos `Lca*`, importa todos en orden y guarda `workbook/LCA.build.xlsm`. `-BuildPanels` ejecuta `ConstruirLca` (hoja **LCA** con todas las bandas P04–P14, P13 y P14); `-InPlace` sobrescribe `LCA.xlsm`. |
| `Test-Workbook.ps1` | Ejecuta `LcaVerificarTodoTexto` (P04–P14, P13, P14) en Excel oculto sobre una copia temporal e imprime PASS/FAIL por suite. Código de salida 0 solo si todo pasa. / Runs every suite headless; exit code 0 only when all pass. |
| `Export-Modules.ps1` | Exporta los módulos `Lca*` de un libro a `src/vba` para conciliar con `git diff`. |
| `LcaModules.ps1` | Lista y orden de módulos. Añadir aquí cada módulo nuevo. |

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-Workbook.ps1 -BuildPanels
powershell -ExecutionPolicy Bypass -File scripts\Test-Workbook.ps1
powershell -ExecutionPolicy Bypass -File scripts\Export-Modules.ps1
```

Las macros de prueba muestran un `MsgBox` cuando se ejecutan a mano con `Alt+F8`; dentro de `VerificarTodo` y de `Test-Workbook.ps1` informan en silencio mediante `LcaReport`. El `.xlsm` es binario: un integrador a la vez.
