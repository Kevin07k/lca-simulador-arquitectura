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

| Tarea / Task | Guía / Guide | Prueba / Test macro | Bloque en la hoja LCA / LCA block |
| --- | --- | --- | --- |
| P03 libro VBA | [P03_VERIFICATION.md](P03_VERIFICATION.md) | `VerificarInstalacion` | portada (col. 1, filas 1–26) |
| P04 CPU / ALU | [P04_CPU_ALU.md](P04_CPU_ALU.md), [P04_INTERFACE.md](P04_INTERFACE.md), [P04_DIAGRAM.md](P04_DIAGRAM.md) | `VerificarCpuAlu` | col. 1, filas 97–211 |
| P05 buses | [P05_BUS_CONTROL.md](P05_BUS_CONTROL.md) | `VerificarBusesP05` | col. 3, filas 75–146 |
| P06 RAM | [P06_RAM.md](P06_RAM.md) | `VerificarRamP06` | col. 3, filas 75–146 |
| P07 tabla de páginas | [P07_MMU.md](P07_MMU.md) | `VerificarMmuP07` | col. 4, filas 1–33 |
| P08 page fault + swap | [P08_PAGE_FAULT_SWAP.md](P08_PAGE_FAULT_SWAP.md) | `VerificarPagerP08` | col. 4, filas 35–104 |
| P09 cachés L1I/L1D/L2 | [P09_CACHE.md](P09_CACHE.md) | `VerificarCachesP09` | col. 4, filas 106–148 |
| P10 caché + paginación | [P10_CACHE_PAGING.md](P10_CACHE_PAGING.md) | `VerificarCachePaginacionP10` | col. 2, filas 100–146 |
| P11 editor + ensamblador + cargador | [P11_ASSEMBLER_LOADER.md](P11_ASSEMBLER_LOADER.md) | `VerificarEnsambladorP11` | col. 2, filas 1–98 |
| **P12 maquina completa (vista de defensa)** | [P12_CPU_CYCLE.md](P12_CPU_CYCLE.md) | `VerificarCicloP12` | **col. 1, filas 28–95** |
| P13 teclado + cola FIFO | [P13_KEYBOARD.md](P13_KEYBOARD.md) | `VerificarTecladoP13` | col. 3, filas 1–36 |
| P14 LCD 8x1 + video | [P14_LCD.md](P14_LCD.md) | `VerificarLcdP14` | col. 3, filas 38–73 |
| **P15 demos aritméticos (datos en tránsito)** | [P15_ARITHMETIC_DEMO.md](P15_ARITHMETIC_DEMO.md) | `VerificarDemoAritmeticaP15` | col. 2, filas 151–238 |
| P16 eco teclado → LCD | [P16_ECHO_DEMO.md](P16_ECHO_DEMO.md) | `VerificarEcoP16` | col. 3, filas 148–202 |
| **P17 memoria: faults y caché** | [P17_MEMORY_DEMO.md](P17_MEMORY_DEMO.md) | `VerificarMemoriaP17` | col. 4, filas 150–217 |
| **P18 controles + sincronización** | [P18_CONTROLS.md](P18_CONTROLS.md) | `VerificarControlesP18` | barra de la máquina, col. 1, filas 28–40 |

`VerificarTodo` (Alt+F8) ejecuta P04–P18 y la prueba de interfaz `VerificarInterfazUI`, muestra un resumen y lo guarda en la hoja **Logs**; `scripts/Test-Workbook.ps1` hace lo mismo sin abrir Excel en pantalla.
`VerificarTodo` runs P04–P18 plus the `VerificarInterfazUI` interface test, shows one summary and stores it on the **Logs** sheet; `scripts/Test-Workbook.ps1` does the same headless.

Capturas / screenshots: [`img/`](img/).
