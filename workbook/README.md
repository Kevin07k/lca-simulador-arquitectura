# Libro / Workbook

`LCA.xlsm` es el simulador ejecutable. Abrirlo en Microsoft Excel de escritorio y habilitar macros solo si se confía en este código revisado. Nunca renombrar un `.xlsx` a `.xlsm`.

`LCA.xlsm` is the executable simulator. Open it in desktop Excel and enable macros only if you trust this reviewed code.

## Hojas / Sheets

| Hoja / Sheet | Contenido / Content |
| --- | --- |
| `Inicio` | Portada. `VerificarInstalacion` escribe `VBA OK - P03` en `B13`. |
| `LCA` | Tablero 2D: todos los bloques lado a lado y uno debajo de otro / 2-D dashboard with every block. |

## Mapa de la hoja LCA / LCA sheet map

El libro se abre en la **máquina P12** (arriba a la izquierda). Cada bloque tiene botones de navegación y un botón **< MAQUINA** para volver; **VISTA GENERAL** aleja el zoom para ver todo el tablero.

| Columna 1 (A–T) | Columna 2 (U–AN) | Columna 3 (AO–BH) | Columna 4 (BI–CB) |
| --- | --- | --- | --- |
| **P12 máquina completa** (filas 1–68) | P11 editor ASM (1–98) | P13 teclado (1–36) | P07 tabla de páginas (1–33) |
| P04 CPU / ALU (70–113) | P10 caché + paginación (100–146) | P14 LCD 8x1 (38–73) | P08 page fault + swap (35–69) |
| P04 diagrama del CPU (115–156) | | P05/P06 buses y RAM (75–108) | P09 cachés L1I/L1D/L2 (71–113) |

Guías: [P04](../docs/P04_INTERFACE.md), [P05](../docs/P05_BUS_CONTROL.md), [P06](../docs/P06_RAM.md), [P07](../docs/P07_MMU.md), [P08](../docs/P08_PAGE_FAULT_SWAP.md), [P09](../docs/P09_CACHE.md), [P10](../docs/P10_CACHE_PAGING.md), [P11](../docs/P11_ASSEMBLER_LOADER.md), [P12](../docs/P12_CPU_CYCLE.md), [P13](../docs/P13_KEYBOARD.md), [P14](../docs/P14_LCD.md).

`ConstruirLca` (Alt+F8) reconstruye todo el tablero.

## Mantener libro y fuente alineados / Keep workbook and source aligned

La fuente de verdad del código es `src/vba`. Para regenerar el libro: `scripts\Build-Workbook.ps1 -BuildPanels -InPlace`; para comprobarlo: `scripts\Test-Workbook.ps1` o `VerificarTodo`. Procedimiento manual: [VBA_WORKFLOW.md](../docs/VBA_WORKFLOW.md). `LCA.build.xlsm` es una salida local ignorada por Git.

The source of truth for code is `src/vba`. Rebuild with `Build-Workbook.ps1 -BuildPanels -InPlace` and verify with `Test-Workbook.ps1` or `VerificarTodo`.
