# P17 — Page fault and cache demo / Programa de page faults y caché

Issue [#17](https://github.com/Kevin07k/lca-simulador-arquitectura/issues/17). Depends on P10 (#10), P11 (#11) and P12 (#12). Platform: Excel `.xlsm` with VBA.

<details open>
<summary><strong>English</strong></summary>

## What it is

A real assembly program, [`programs/memoria/memoria.asm`](../programs/memoria/memoria.asm), runs on the P12 CPU through the whole P10 path (MMU → CLOCK pager → L1I/L1D → L2 → RAM). The P17 block of the **LCA** sheet (column 4, rows 150–217, below the P09 caches; machine bar button **MEMORIA P17**, macro `LcaIrMemoria`) draws the 7 RAM frames, the 15 virtual pages, every cache line and one table row per instruction, and animates each page and cache line that moves.

| Phase | Instructions | What it shows |
|---|---|---|
| 1 DISPERSION | seven `ST` to pages 8, 10, 12, 14, 9, 11, 13 | every store is a page fault **and** an L1D miss |
| 2 REPETICION | `LD BX,[100h]` four times (loop) | pass 1 recovers the modified byte from the swap; passes 2–4 **hit** L1D |
| 3 RESIDENTE | `LD DX,[1A0h]` | no page fault (page 13 is in RAM) but still an L1D miss |
| 4 COPIA | `ST [101h],BX`, `HLT` | the recovered byte is stored again |

The program has 21 instructions (84 bytes, 3 code pages) and runs **30** instructions. Code pages 0–2 plus data pages 8–14 are **10 pages for 7 pageable frames**, so the CLOCK must evict.

## Expected counters (derived by hand)

| Counter | Value | Why |
|---|---|---|
| Page faults | **11** | 10 pages touched once + page 8 faults again in phase 2 |
| Evictions | **4** | faults − 7 frames |
| Write-backs | **3** | victims: code page 0 (clean), then dirty pages 8, 10, 12 |
| L1D hits / misses | **3 / 10** | loop passes 2–4 hit; misses = 7 stores + first `LD` + `LD [1A0h]` + `ST [101h]` |
| L1I hits / misses | **93 / 27** | 30 fetches × 4 bytes = 120 accesses; 21 cold lines + 6 loop conflicts (`LD` and `JNZ` share one of the 2 L1I lines, 3 passes × 2) |
| L2 hits / misses | **6 / 31** | L2 is asked once per L1 miss (27 + 10 = 37): the 6 loop refills hit, the 31 cold lines (21 code + 10 data) miss |
| BX | **A1h** | page 8 was evicted dirty, written back and read again |

Observed values equal the expected ones; the block shows both with an **OK** column, and `VerificarMemoriaP17` asserts them (97 checks).

## Access sequence

| # | PC | Instruction | Data | Page fault | Victim | L1D |
|---|---|---|---|---|---|---|
| 1–2 | 000–004 | `MOV AX,#0A1h`, `ST [100h],AX` | 100h | code 0, data 8 | | miss |
| 3–10 | 008–024 | five more `MOV`/`ST` pairs (140h, 180h, 1C0h, 120h) | | code 1, data 10, 12, 14, 9 | | miss |
| 11–12 | 028–02C | `MOV`, `ST [160h]` | 160h | page 11 | VPN 0 (clean) | miss |
| 13–14 | 030–034 | `MOV`, `ST [1A0h]` | 1A0h | page 13 | VPN 8, **write-back** | miss |
| 15–16 | 038–03C | `MOV CX,#04h`, `LD BX,[100h]` | 100h | page 8 again | VPN 10, **write-back** | miss |
| 17 | 040 | `DEC CX` | | code page 2 | VPN 12, **write-back** | |
| 18–27 | 044–03C… | `JNZ`, then 3 more passes | 100h | none | | **hit** ×3 |
| 28 | 048 | `LD DX,[1A0h]` | 1A0h | none | | miss |
| 29–30 | 04C–050 | `ST [101h],BX`, `HLT` | 101h | none | | miss |

## How to reproduce

1. Open `workbook/LCA.xlsm`, press **MEMORIA P17** on the machine bar.
2. **CARGAR PROGRAMA**, then **PASO** to watch each page fault (the page flies from the swap row to its frame; a dirty victim flies back), or **EJECUTAR** for the whole run.
3. Compare the **ESPERADO / OBSERVADO** table; every row shows **OK** when the program halts.
4. `VerificarMemoriaP17` (or `VerificarTodo`) checks it without the interface.

Screenshots: [`img/P17_mid.png`](img/P17_mid.png) (first 16 instructions), [`img/P17_end.png`](img/P17_end.png) (finished).

## Files

`src/vba/LcaMemDemo.bas` (program, access log, expected counters), `LcaMemPanel.bas` (block), `LcaMemTests.bas` (`VerificarMemoriaP17`), `programs/memoria/memoria.asm` (checked against the workbook by `scripts/Test-Workbook.ps1`).

</details>

<details>
<summary><strong>Español</strong></summary>

## Qué es

Un programa en ensamblador, [`programs/memoria/memoria.asm`](../programs/memoria/memoria.asm), corre en la CPU de P12 a través de todo el camino de P10 (MMU → paginador CLOCK → L1I/L1D → L2 → RAM). El bloque P17 de la hoja **LCA** (columna 4, filas 150–217, bajo las cachés de P09; botón **MEMORIA P17** en la barra de la máquina, macro `LcaIrMemoria`) dibuja los 7 marcos de RAM, las 15 páginas virtuales, cada línea de caché y una fila de tabla por instrucción, y anima cada página y línea que se mueve.

1. **Dispersión**: siete `ST` a páginas distintas; cada uno es una falta de página y un fallo de L1D.
2. **Repetición**: `LD BX,[100h]` cuatro veces; la primera recupera del swap el byte modificado y las otras tres acierran en L1D.
3. **Residente**: `LD DX,[1A0h]` no produce falta de página, pero sí un fallo de caché (página residente ≠ línea en caché).
4. **Copia**: el dato recuperado se guarda en `[101h]` y `HLT`.

Con 3 páginas de código y 7 de datos hay **10 páginas para 7 marcos**: la CLOCK desaloja. Contadores esperados (deducidos a mano, tabla de arriba): 30 instrucciones, 11 faltas, 4 evicciones, 3 write-backs, L1D 3/10, L1I 93/27, L2 6/31 y BX = A1h. El bloque muestra esperado y observado con una columna **OK**; `VerificarMemoriaP17` los comprueba (97 verificaciones).

## Cómo reproducir

1. Abrir `workbook/LCA.xlsm` y pulsar **MEMORIA P17**.
2. **CARGAR PROGRAMA**, luego **PASO** (cada falta de página se anima) o **EJECUTAR**.
3. Comparar la tabla **ESPERADO / OBSERVADO**.

</details>
