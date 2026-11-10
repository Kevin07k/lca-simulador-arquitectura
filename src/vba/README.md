# Fuente VBA / VBA source

Módulos estándar exportados de `workbook/LCA.xlsm`, en ASCII con CRLF. El orden de importación está en [`scripts/LcaModules.ps1`](../../scripts/LcaModules.ps1); `scripts/Build-Workbook.ps1` los importa todos. Los módulos de **motor** no tocan hojas; los de **panel** dibujan una banda de la hoja única **LCA**; los de **prueba** informan con `LcaReport`.

Standard modules exported from `workbook/LCA.xlsm` (ASCII, CRLF). **Engine** modules never touch worksheets; **panel** modules draw one band of the single **LCA** sheet; **test** modules report through `LcaReport`.

| Tarea / Task | Motor / Engine | Panel | Prueba / Test | Guía / Guide |
| --- | --- | --- | --- | --- |
| P03 | `LcaInicio` (`VerificarInstalacion`, `LcaReport`, `VerificarTodo`) | — | — | [VBA_WORKFLOW](../../docs/VBA_WORKFLOW.md) |
| P04 | `LcaCpu`, `LcaAlu` | `LcaPanel`, `LcaDiagrama` | `LcaCpuTests` → `VerificarCpuAlu` | [P04](../../docs/P04_CPU_ALU.md) |
| P05 | `LcaBus`, `LcaControl` | `LcaBusPanel` | `LcaBusTests` → `VerificarBusesP05` | [P05](../../docs/P05_BUS_CONTROL.md) |
| P06 | `LcaRam`, `LcaMemoryPath` | `LcaBusPanel` | `LcaRamTests` → `VerificarRamP06` | [P06](../../docs/P06_RAM.md) |
| P07 | `LcaMmu`, `LcaMmuPath` | `LcaMmuPanel` | `LcaMmuTests` → `VerificarMmuP07` | [P07](../../docs/P07_MMU.md) |
| P08 | `LcaSwap`, `LcaPager` | `LcaPagerPanel` | `LcaPagerTests` → `VerificarPagerP08` | [P08](../../docs/P08_PAGE_FAULT_SWAP.md) |
| P09 | `LcaCache` | `LcaCachePanel` | `LcaCacheTests` → `VerificarCachesP09` | [P09](../../docs/P09_CACHE.md) |
| P10 | `LcaCachePath` | `LcaCachePathPanel` | `LcaCachePathTests` → `VerificarCachePaginacionP10` | [P10](../../docs/P10_CACHE_PAGING.md) |
| P11 | `LcaAssembler`, `LcaLoader` | `LcaEditorPanel` | `LcaAssemblerTests` → `VerificarEnsambladorP11` | [P11](../../docs/P11_ASSEMBLER_LOADER.md) |
| P13 | `LcaKeyboard` | `LcaKeyboardPanel` | `LcaKeyboardTests` → `VerificarTecladoP13` | [P13](../../docs/P13_KEYBOARD.md) |
| P14 | `LcaLcd` | `LcaLcdPanel` | `LcaLcdTests` → `VerificarLcdP14` | [P14](../../docs/P14_LCD.md) |
| Hoja / sheet | `LcaLayout` (`ConstruirLca`, `LcaRows`, navegación `LcaIr*`) | | | |

## Recorrido de memoria / Memory path

```
CPU ── LcaCachePath (P10) ── LcaPager (P08) ── LcaMmu (P07)      residencia / residency
             │                    └─ invalida marco / invalidates frame ─┐
             └── LcaCache (P09: L1I, L1D, L2) ── LcaRam (P06) ◄──────────┘
                     └── LcaControl + LcaBus (P05): ciclos FILL / LD / ST
```

`LcaMmuPath` (P07) y `PagerFetchFour` / `PagerReadByte` / `PagerStoreByte` (P08) siguen siendo recorridos **sin caché** para sus propias pruebas. P12 debe usar `LcaCachePath`.
`LcaMmuPath` (P07) and the P08 `Pager*` byte paths remain **uncached** paths for their own tests; P12 must use `LcaCachePath`.

## Hoja LCA / LCA sheet

`ConstruirLca` borra la hoja y construye las bandas en orden: CPU (1–44), diagrama (46–87), buses/RAM (89–122), P07 (124–156), P08 (158–192), P09 (194–236), P10 (238–284), P11 (286–383), P13 (385–420), P14 (421–456). Cada `Construir...` reconstruye solo su banda; las direcciones se desplazan con `LcaRows` según `ROW_BASE`.

`ConstruirLca` clears the sheet and builds every band in that order; each `Construir...` macro rebuilds only its own band.

## Convenciones / Conventions

- Errores `vbObjectError + n`, un bloque por tarea: CPU 2001–2005 (prueba 2100), bus 2051–2054, RAM 2061–2064, MMU 2071–2076, swap/pager 2081–2083, caché 2091–2092, P10 2101–2102, P11 2111–2113, teclado P13 2131–2132, LCD P14 2141–2142.
- Estado privado en el módulo y acceso por funciones; las pruebas reinician el estado que usan.
- Solo ASCII en `.bas` (Excel importa en ANSI). No fusionar el `.xlsm` como texto.
