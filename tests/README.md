# Verificación / Verification

Las pruebas viven en el libro como macros (`src/vba/Lca*Tests.bas`). No hay framework externo: cada macro reinicia el estado que usa, comprueba resultados exactos e informa PASS/FAIL con el número de comprobaciones.

The tests are workbook macros (`src/vba/Lca*Tests.bas`). Each one resets the state it uses, checks exact results and reports PASS/FAIL with its assertion count.

## Ejecutar todo / Run everything

- En Excel: `Alt+F8` → `VerificarTodo`. Un solo mensaje con una línea por suite (P04–P16 y la prueba de interfaz); el reporte también queda en la hoja **Logs** y en la portada.
- Sin abrir Excel en pantalla / headless (Windows + Excel de escritorio):

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-Workbook.ps1 -BuildPanels
powershell -ExecutionPolicy Bypass -File scripts\Test-Workbook.ps1
```

`Test-Workbook.ps1` trabaja sobre una copia temporal y termina con código 0 solo si todo pasa.

## Suites

| Macro | Tarea / Task | Último resultado observado / Last observed (2026-11-10) |
| --- | --- | --- |
| `VerificarCpuAlu` | P04 | PASS — 197120 casos, 1379895 comprobaciones |
| `VerificarBusesP05` | P05 | PASS — 38 |
| `VerificarRamP06` | P06 | PASS — 38 |
| `VerificarMmuP07` | P07 | PASS — 86 |
| `VerificarPagerP08` | P08 | PASS — 32 |
| `VerificarCachesP09` | P09 | PASS — 77 |
| `VerificarCachePaginacionP10` | P10 | PASS — 58 |
| `VerificarEnsambladorP11` | P11 | PASS — 280 |
| `VerificarCicloP12` | P12 | PASS — 45 |
| `VerificarTecladoP13` | P13 | PASS — 94 |
| `VerificarLcdP14` | P14 | PASS — 52 |
| `VerificarDemoAritmeticaP15` | P15 | PASS — 153 (cinco programas: resultados, pasos, banderas, memoria, fallos de página, traza por vuelta, STEP = RUN) |
| `VerificarEcoP16` | P16 | PASS — 67 (eco IN/OUT: orden FIFO, espera y continuación, ocho caracteres, noveno, cola llena, fin con `.`, STEP = RUN) |
| `VerificarInterfazUI` | Interfaz | PASS — 46 (cuadricula 2D, hojas LCA + Logs, elementos de cada bloque, refresco de vistas) |

Cada macro reinicia CPU, memoria o cachés; no ejecutarlas durante una simulación que se quiera conservar. Evidencia inicial de P03: [P03_VERIFICATION.md](../docs/P03_VERIFICATION.md).

Each macro resets CPU, memory or caches; do not run them during a simulation you want to keep.

`scripts/Test-Workbook.ps1` además comprueba que `programs/arith/*.asm` y `programs/echo/eco.asm` coinciden con las fuentes del libro (`LcaArith.bas`, `LcaEcho.bas`) / also checks that `programs/arith/*.asm` and `programs/echo/eco.asm` match the workbook source.
