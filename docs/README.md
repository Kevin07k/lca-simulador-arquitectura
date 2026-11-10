# Documentación / Documentation

Índice de guías por tarea. Cada guía es bilingüe (English / Español) salvo que se indique.
Index of task guides. Each guide is bilingual unless noted.

## Contratos y flujo / Contracts and workflow

| Documento / Document | Contenido / Content |
| --- | --- |
| [ARCHITECTURE.md](ARCHITECTURE.md) · [ARQUITECTURA_ES.md](ARQUITECTURA_ES.md) | P02: ISA, mapa de memoria, paginación, cachés, buses, E/S y errores. |
| [PROJECT_MANAGEMENT.md](PROJECT_MANAGEMENT.md) | Ramas, PR, revisión y tablero / branches, PRs, review and board. |
| [VBA_WORKFLOW.md](VBA_WORKFLOW.md) | Abrir el libro e importar/exportar módulos / open the workbook and sync modules. |
| [IDENTITY.md](IDENTITY.md) | Nombre, logo y colores / name, logo and colours. |

## Guías por tarea / Task guides

| Tarea / Task | Guía / Guide | Prueba / Test macro | Banda en hoja LCA / LCA band |
| --- | --- | --- | --- |
| P03 libro VBA | [P03_VERIFICATION.md](P03_VERIFICATION.md) | `VerificarInstalacion` | hoja `Inicio` |
| P04 CPU / ALU | [P04_CPU_ALU.md](P04_CPU_ALU.md), [P04_INTERFACE.md](P04_INTERFACE.md), [P04_DIAGRAM.md](P04_DIAGRAM.md) | `VerificarCpuAlu` | filas 1–87 |
| P05 buses | [P05_BUS_CONTROL.md](P05_BUS_CONTROL.md) | `VerificarBusesP05` | filas 89–122 |
| P06 RAM | [P06_RAM.md](P06_RAM.md) | `VerificarRamP06` | filas 89–122 |
| P07 tabla de páginas | [P07_MMU.md](P07_MMU.md) | `VerificarMmuP07` | filas 124–156 |
| P08 page fault + swap | [P08_PAGE_FAULT_SWAP.md](P08_PAGE_FAULT_SWAP.md) | `VerificarPagerP08` | filas 158–192 |
| P09 cachés L1I/L1D/L2 | [P09_CACHE.md](P09_CACHE.md) | `VerificarCachesP09` | filas 194–236 |
| P10 caché + paginación | [P10_CACHE_PAGING.md](P10_CACHE_PAGING.md) | `VerificarCachePaginacionP10` | filas 238–284 |

`VerificarTodo` (Alt+F8) ejecuta P04–P10 y muestra un resumen; `scripts/Test-Workbook.ps1` hace lo mismo sin abrir Excel en pantalla.
`VerificarTodo` runs P04–P10 and shows one summary; `scripts/Test-Workbook.ps1` does the same headless.

Capturas / screenshots: [`img/`](img/).
