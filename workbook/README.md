# Libro / Workbook

`LCA.xlsm` es el simulador ejecutable. Abrirlo en Microsoft Excel de escritorio y habilitar macros solo si se confía en este código revisado. Nunca renombrar un `.xlsx` a `.xlsm`.

`LCA.xlsm` is the executable simulator. Open it in desktop Excel and enable macros only if you trust this reviewed code.

## Hojas / Sheets

| Hoja / Sheet | Contenido / Content |
| --- | --- |
| `LCA` | Tablero 2D con todo el simulador: portada, máquina animada y un bloque gráfico por módulo. / 2-D dashboard with the whole simulator. |
| `Logs` | Trazas completas de cada módulo (la más reciente arriba) y el último `VerificarTodo`. / Full traces and the last test report. |

## Mapa de la hoja LCA / LCA sheet map

El libro se abre en la **portada** (arriba a la izquierda): 10 tarjetas, una por cada parte del simulador. Sus tarjetas, la barra de navegación de la máquina y el botón **< MAQUINA** de cada bloque llevan a cualquier parte; **VISTA GENERAL** aleja el zoom para ver todo el tablero.

| Columna 1 (A–T) | Columna 2 (U–AN) | Columna 3 (AO–BH) | Columna 4 (BI–CB) |
| --- | --- | --- | --- |
| **Portada** (filas 1–26) | P11 editor ASM (1–98) | P13 teclado (1–36) | P07 tabla de páginas + traducción bit a bit (1–33) |
| **P12 máquina animada** (28–95) | P10 caché + paginación (100–146) | P14 LCD 8x1 (38–73) | P08 swap + reloj CLOCK + mapa de la RAM (35–104) |
| P04 CPU + bits + gráfica (97–168) | **P15 demos aritméticos: datos en tránsito** (151–238) | P05/P06 buses + diagrama tri-estado (75–146) | P09 cachés + gráfica hit/miss (106–148) |
| P04 diagrama del CPU (170–211) | | | |

Guías: [P04](../docs/P04_INTERFACE.md), [P05](../docs/P05_BUS_CONTROL.md), [P06](../docs/P06_RAM.md), [P07](../docs/P07_MMU.md), [P08](../docs/P08_PAGE_FAULT_SWAP.md), [P09](../docs/P09_CACHE.md), [P10](../docs/P10_CACHE_PAGING.md), [P11](../docs/P11_ASSEMBLER_LOADER.md), [P12](../docs/P12_CPU_CYCLE.md), [P13](../docs/P13_KEYBOARD.md), [P14](../docs/P14_LCD.md), [P15](../docs/P15_ARITHMETIC_DEMO.md).

`ConstruirLca` (Alt+F8) reconstruye el tablero y la hoja Logs. `VerificarInstalacion` escribe `VBA OK - P03` en la línea de estado de la portada.

## Mantener libro y fuente alineados / Keep workbook and source aligned

La fuente de verdad del código es `src/vba`. Para regenerar el libro: `scripts\Build-Workbook.ps1 -BuildPanels -InPlace` (también incrusta el logo de `assets/lca-chip.png` en la portada); para comprobarlo: `scripts\Test-Workbook.ps1` o `VerificarTodo`. Procedimiento manual: [VBA_WORKFLOW.md](../docs/VBA_WORKFLOW.md). `LCA.build.xlsm` es una salida local ignorada por Git.

The source of truth for code is `src/vba`. Rebuild with `Build-Workbook.ps1 -BuildPanels -InPlace` and verify with `Test-Workbook.ps1` or `VerificarTodo`.
