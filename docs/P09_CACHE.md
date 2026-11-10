# P09 — L1I, L1D and L2 caches / Cachés L1I, L1D y L2

<details open>
<summary><strong>English</strong></summary>

## What changed

`src/vba/LcaCache.bas` implements the three caches from the P02 contract (section 4 of [ARCHITECTURE.md](ARCHITECTURE.md)). It has no worksheet dependency and reaches RAM only through `RamReadUser` / `RamWriteUser`, so I/O (`E0..EFh`), the reserved frame (`F0..FFh`) and out-of-range addresses are rejected before any state changes.

| Level | Size | Lines | Notes |
| --- | --- | --- | --- |
| L1I | 8 B | 2 | Instruction fetches (`LCA_CACHE_FETCH`) |
| L1D | 8 B | 2 | Data loads/stores (`LCA_CACHE_DATA`) |
| L2 | 32 B | 8 | Unified, consulted on every L1 miss |

## Policy (documented as required)

- Tags are **physical**. `line = PA \ 4`, `index = line Mod lines`, `tag = line \ lines`. Lines never cross a 32-byte frame.
- Direct mapped, replacement by overwrite. An L1 miss queries L2; an L2 miss reads the whole 4-byte line from RAM (`CacheLastFillPA` reports the base address so the bus layer can emit the four READ cycles; a hit creates no external cycle).
- Counters change once per **CPU-requested byte**, never per fill byte. L1I and L1D keep separate hit/miss; L2 counts one query per L1 miss.
- Writes (`CacheWrite`) are **write-through, write-allocate**: on an L1D miss the line is loaded first, then L1D, L2 (if resident) and RAM are updated, in that order after validation. L1I is kept coherent without counting an access, so a stored byte is never fetched stale.
- `CacheInvalidateFrame(pfn)` drops every line of that frame in all three caches (used before P08 reuses a PFN); `CacheReset` invalidates everything and zeroes counters, fills and trace.

## API

`CacheRead(PA, kind)`, `CacheWrite(PA, byte)`, `CacheReset`, `CacheInvalidateAll`, `CacheInvalidateFrame(pfn)`, `CacheHits(level)`, `CacheMisses(level)`, `CacheContains(level, PA)`, `CacheLineValid/Tag/Byte`, `CacheRamFills`, `CacheLastFillPA`, `CacheLastEvent`, `CacheTraceCount/Line`. Errors: `LCA_ERR_CACHE_ARGUMENT` (2091) plus the RAM/CPU codes for protection, range and byte errors. `LcaCachePanel.bas` builds the cache band of the single **LCA** sheet (`ConstruirPanelCaches`, or `ConstruirLca` for every band) with per-level hit/miss counters, every line (valid, tag, bytes, covered PA) and the live trace. `LcaCacheTests.bas` runs `VerificarCachesP09`.

## Expected traces (physical addresses)

| Case | Accesses | L1I h/m | L1D h/m | L2 h/m |
| --- | --- | --- | --- | --- |
| P02 fetch, RAM `40..43h = 10 00 2A 00` | fetch 40h, 41h, 42h, 43h | 3/1 | 0/0 | 0/1 |
| Data line | read 6Dh, 6Dh, 6Ch | 0/0 | 2/1 | 0/1 |
| Write hit then read | + write 6Dh=55h, read 6Dh | 0/0 | 4/1 | 0/1 |
| Write-allocate | + write 70h=77h, read 70h | 0/0 | 5/2 | 0/2 |
| L1I conflict | fetch 40h, 48h, 40h | 0/3 | 0/0 | 1/2 |
| L2 conflict | data 00h, 20h, 00h | 0/0 | 0/3 | 0/3 |
| Fetch ≠ data | fetch 40h, data 40h | 0/1 | 0/1 | 1/1 |

Each row after the first two starts from `CacheReset` except the "write hit" and "write-allocate" rows, which continue the "data line" row. The same numbers are asserted by `VerificarCachesP09`.

## Reproduce in desktop Excel

1. Work on a copy of the latest `workbook/LCA.xlsm`. In the VBA editor (`Alt+F11`) import `LcaCache.bas`, `LcaLayout.bas`, `LcaCachePanel.bas`, `LcaCacheTests.bas` (and the updated panel modules `LcaPanel`, `LcaDiagrama`, `LcaBusPanel`; or simply run `scripts/Build-Workbook.ps1`) (the P06 modules `LcaRam`, `LcaCpu` must already be present). **Debug → Compile VBAProject**.
2. `Alt+F8` → `VerificarCachesP09`: expect PASS with the assertion count shown by the macro. Record the **observed** message.
3. `Alt+F8` → `ConstruirPanelCaches`, run `ConstruirLca` and scroll to the cache band (rows 124-166), then click **NEXT ACCESS** eleven times. Expected counters: after the 4th click L1I 3/1, L1D 0/0, L2 0/1; after the 11th click L1I 3/3, L1D 3/2, L2 1/4 (the last two fetches, 48h then 40h, conflict on L1I index 0 and 40h is served by an L2 hit). Take a screenshot of the sheet.
4. Run `VerificarRamP06`, `VerificarBusesP05` and `VerificarCpuAlu` as regressions.
5. Save the copy as `workbook/LCA.xlsm`, export the three new modules, compare with `src/vba` (including `Attribute VB_Name`) and reconcile. Do not merge `.xlsm` as text.

## Limits and pending

P09 works with physical addresses; translation (P07) and the combined path (P10) are not wired here, and no bus cycles are emitted yet (the fill address is exposed for P10/P12). Static review only in this delivery: desktop Excel compilation, macro execution and screenshots are **pending**, not PASS.

</details>

<details>
<summary><strong>Español</strong></summary>

## Qué se implementó

`src/vba/LcaCache.bas` implementa las tres cachés del contrato P02 (sección 4 de [ARQUITECTURA_ES.md](ARQUITECTURA_ES.md)). No depende de hojas y accede a la RAM solo con `RamReadUser` / `RamWriteUser`, de modo que E/S (`E0..EFh`), el marco reservado (`F0..FFh`) y direcciones fuera de rango se rechazan antes de cambiar estado.

| Nivel | Tamaño | Líneas | Notas |
| --- | --- | --- | --- |
| L1I | 8 B | 2 | Fetch de instrucciones (`LCA_CACHE_FETCH`) |
| L1D | 8 B | 2 | Lecturas/escrituras de datos (`LCA_CACHE_DATA`) |
| L2 | 32 B | 8 | Unificada, consultada en cada fallo de L1 |

## Política (documentada)

- Etiquetas **físicas**. `línea = PA \ 4`, `índice = línea Mod líneas`, `etiqueta = línea \ líneas`. Ninguna línea cruza un marco de 32 bytes.
- Mapeo directo y reemplazo por sobrescritura. Un fallo de L1 consulta L2; un fallo de L2 lee la línea completa de 4 bytes de RAM (`CacheLastFillPA` informa la dirección base para que la capa de bus emita las cuatro lecturas; un acierto no genera ciclo externo).
- Los contadores cambian una vez por **byte pedido por la CPU**, no por byte de relleno. L1I y L1D llevan hit/miss separados; L2 cuenta una consulta por fallo de L1.
- Las escrituras (`CacheWrite`) son **write-through con write-allocate**: en fallo de L1D se carga la línea y luego se actualizan L1D, L2 (si está presente) y RAM, tras validar. L1I se mantiene coherente sin contar acceso, así que nunca se lee un byte escrito desactualizado.
- `CacheInvalidateFrame(pfn)` elimina las líneas de ese marco en las tres cachés (antes de que P08 reutilice un PFN); `CacheReset` invalida todo y pone en cero contadores, rellenos y traza.

## Interfaz

`CacheRead(PA, tipo)`, `CacheWrite(PA, byte)`, `CacheReset`, `CacheInvalidateAll`, `CacheInvalidateFrame(pfn)`, `CacheHits(nivel)`, `CacheMisses(nivel)`, `CacheContains(nivel, PA)`, `CacheLineValid/Tag/Byte`, `CacheRamFills`, `CacheLastFillPA`, `CacheLastEvent`, `CacheTraceCount/Line`. Errores: `LCA_ERR_CACHE_ARGUMENT` (2091) y los códigos de RAM/CPU para protección, rango y byte. `LcaCachePanel.bas` construye la banda de cachés de la hoja única **LCA** (`ConstruirPanelCaches`, o `ConstruirLca` para todas) con contadores hit/miss por nivel, todas las líneas (válida, etiqueta, bytes, PA cubiertas) y la traza real. `LcaCacheTests.bas` ejecuta `VerificarCachesP09`.

## Trazas esperadas (direcciones físicas)

Ver la tabla de la versión en inglés: los valores hit/miss son idénticos y son los que verifica `VerificarCachesP09`. Caso P02: fetch 40h–43h con RAM `10 00 2A 00` → L1I 3/1, L2 0/1.

## Reproducir en Excel de escritorio

1. Trabajar sobre una copia del último `workbook/LCA.xlsm`. En el editor VBA (`Alt+F11`) importar `LcaCache.bas`, `LcaLayout.bas`, `LcaCachePanel.bas` y `LcaCacheTests.bas` (y los paneles actualizados `LcaPanel`, `LcaDiagrama`, `LcaBusPanel`; o solo ejecutar `scripts/Build-Workbook.ps1`) (los módulos de P06 `LcaRam` y `LcaCpu` ya deben estar). **Depurar → Compilar VBAProject**.
2. `Alt+F8` → `VerificarCachesP09`: se espera PASS con el número de comprobaciones que muestre la macro. Registrar el mensaje **observado**.
3. `Alt+F8` → `ConstruirPanelCaches` y pulsar **SIGUIENTE ACCESO** once veces; Esperado: tras el 4.º clic L1I 3/1, L1D 0/0, L2 0/1; tras el 11.º L1I 3/3, L1D 3/2, L2 1/4 (los dos últimos fetch, 48h y 40h, chocan en el índice 0 de L1I y 40h se sirve con un acierto de L2). Tomar captura de la hoja.
4. Ejecutar `VerificarRamP06`, `VerificarBusesP05` y `VerificarCpuAlu` como regresión.
5. Guardar la copia como `workbook/LCA.xlsm`, exportar los tres módulos nuevos, compararlos con `src/vba` (incluido `Attribute VB_Name`) y conciliar. No fusionar `.xlsm` como texto.

## Límites y pendientes

P09 trabaja con direcciones físicas; la traducción (P07) y el recorrido combinado (P10) no están conectados aquí y aún no se emiten ciclos de bus (la dirección de relleno queda expuesta para P10/P12). Esta entrega solo se revisó de forma estática: compilación en Excel de escritorio, ejecución de macros y capturas están **pendientes**, no PASS.

</details>
