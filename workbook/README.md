# Libro / Workbook

`LCA.xlsm` es el simulador ejecutable. Abrirlo en Microsoft Excel de escritorio y habilitar macros solo si se confía en este código revisado. Nunca renombrar un `.xlsx` a `.xlsm`.

`LCA.xlsm` is the executable simulator. Open it in desktop Excel and enable macros only if you trust this reviewed code.

## Hojas / Sheets

| Hoja / Sheet | Contenido / Content |
| --- | --- |
| `Inicio` | Portada. `VerificarInstalacion` escribe `VBA OK - P03` en `B13`. |
| `LCA` | Todas las bandas en una sola hoja, de arriba abajo / every band on one sheet, top to bottom. |

| Filas / Rows | Banda / Band | Guía / Guide |
| --- | --- | --- |
| 1–44 | P04 CPU / ALU: registros, ALU, flags, historial | [P04_INTERFACE](../docs/P04_INTERFACE.md) |
| 46–87 | P04 diagrama de bloques del CPU | [P04_DIAGRAM](../docs/P04_DIAGRAM.md) |
| 89–122 | P05/P06 buses, micropalabra y RAM | [P05](../docs/P05_BUS_CONTROL.md), [P06](../docs/P06_RAM.md) |
| 124–156 | P07 tabla de páginas y traducción | [P07](../docs/P07_MMU.md) |
| 158–192 | P08 page fault, CLOCK y swap | [P08](../docs/P08_PAGE_FAULT_SWAP.md) |
| 194–236 | P09 cachés L1I, L1D, L2 | [P09](../docs/P09_CACHE.md) |
| 238–284 | P10 caché + paginación | [P10](../docs/P10_CACHE_PAGING.md) |

`ConstruirLca` (Alt+F8) reconstruye la hoja completa. Las bandas tienen botones de navegación (`VER P08`, `VER P09`, `DIAGRAMA`...).

## Mantener libro y fuente alineados / Keep workbook and source aligned

La fuente de verdad del código es `src/vba`. Para regenerar el libro: `scripts\Build-Workbook.ps1 -BuildPanels -InPlace`; para comprobarlo: `scripts\Test-Workbook.ps1` o `VerificarTodo`. Procedimiento manual: [VBA_WORKFLOW.md](../docs/VBA_WORKFLOW.md). `LCA.build.xlsm` es una salida local ignorada por Git.

The source of truth for code is `src/vba`. Rebuild with `Build-Workbook.ps1 -BuildPanels -InPlace` and verify with `Test-Workbook.ps1` or `VerificarTodo`.
