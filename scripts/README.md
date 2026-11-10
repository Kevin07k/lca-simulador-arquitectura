# Scripts de automatización / Automation scripts

Requisitos / Requirements: Windows, Excel de escritorio, PowerShell. Una sola vez, en Excel:
*Archivo > Opciones > Centro de confianza > Configuración > Configuración de macros > "Confiar en el acceso al modelo de objetos de proyectos de VBA"*. Los scripts no cambian esta opción.

| Script | Qué hace / What it does |
| --- | --- |
| `Build-Workbook.ps1` | Recrea el libro desde `src/vba`: copia `workbook/LCA.xlsm`, quita los módulos `Lca*`, importa todos en orden y guarda `workbook/LCA.build.xlsm`. `-BuildPanels` ejecuta `ConstruirLca`: una sola hoja **LCA** con las bandas de CPU, diagrama, buses y cachés; `-InPlace` sobrescribe `LCA.xlsm`. |
| `Export-Modules.ps1` | Exporta los módulos `Lca*` de un libro a `src/vba` para conciliar con `git diff`. |
| `LcaModules.ps1` | Lista y orden de módulos. Añadir aquí cada módulo nuevo (P07, P08...). |

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-Workbook.ps1 -BuildPanels
powershell -ExecutionPolicy Bypass -File scripts\Export-Modules.ps1
```

Las macros de prueba (`VerificarCachesP09`, etc.) muestran un `MsgBox`, por eso se ejecutan a mano con `Alt+F8`. El `.xlsm` es binario: un integrador a la vez.
