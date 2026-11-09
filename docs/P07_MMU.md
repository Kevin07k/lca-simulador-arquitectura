# P07 — Page table and translation / Tabla de páginas y traducción

## Implemented / Implementado

`LcaMmu.bas` owns a table of **16 virtual pages**, each 32 bytes (`000–1FFh`, 512 bytes). Entries expose `VPN`, `Present`, `PFN`, `Dirty`, `Referenced`, `Reserved`. VPN 0–14 start absent with **PFN = -1**, and VPN 15 (`1E0–1FFh`) is permanently reserved. PFN 0–6 are allocatable; PFN 7 (`E0–FFh`) is never mapped to a user page. The table rejects duplicate ownership of a physical frame.

`LcaMmuPath.bas` joins translation to the P06 RAM/P05 bus path. It prevalidates four fetch byte translations, then clocks each RAM read and internal latch. `MmuStoreByte` transfers register→MDR→RAM and marks `dirty/ref` **after** the successful bus/RAM operation. `LcaMmuPanel.bas` displays all 16 entries in the `P07 Paginas` worksheet, permits manual translation and runs one real fetch/store demonstration. `LcaMmuTests.bas` verifies boundaries, permissions, faults, flags, frame ownership and bus integration.

`LcaMmu` mantiene 16 entradas de 32 bytes: 15 páginas de código/datos asignables y una página reservada de E/S. Una página ausente tiene `P=0, PFN=-1`; una página presente en el marco cero tiene `P=1, PFN=0`. Son estados diferentes. P08 atenderá `PAGE_FAULT`, reemplazo CLOCK y swap. P09 añadirá cachés físicas y P12 decodificará y ejecutará las instrucciones.

## Contract / Contrato

| Operation / Operación | Behavior / Resultado |
| --- | --- |
| `MmuReset` | PTE[0..15] absent, PFN=-1, dirty=ref=0; VPN 15 reserved; frame owners 0..6 free. Does not reset RAM/CPU/buses. |
| `MmuVpnOf(VA)`, `MmuOffsetOf(VA)` | Validate `0..511`; `VPN=VA\32`, `offset=VA Mod 32`. These split even a reserved VA such as 511 without granting access. |
| `MmuMap(VPN,PFN)` | Only VPN 0..14 and PFN 0..6, one resident VPN per PFN. P08 will call after loading a frame/invalidation. |
| `MmuTryTranslate(VA,kind,ByRef PA)` | Validates range and permissions. Present → `True`, PA `PFN*32+offset`. Absent → `False`, PA=-1, `MmuLastFaultVPN/VA`. Invalid region → fatal error. |
| `MmuTranslate(VA,kind)` | Same validation; absent raises recoverable `LCA_ERR_PAGE_FAULT` (`PAGE_FAULT VPN=n`). P08 catches, loads and retries. |
| `MmuCommitAccess(VA,kind,PA)` | Verifies the mapping still equals PA; sets `ref=1`, and `dirty=1` only on completed WRITE. Translation alone changes neither bit. |
| `MmuUnmap(VPN)` | Releases a clean resident page; rejects dirty pages. P08 must write back dirty bytes, invalidate caches and clear dirty first. |
| `MmuFetchFour(VA)`, `MmuStoreByte(VA,R)` | P07 adapter around P06/P05. Fetch leaves PC unchanged; P12 owns decode, PC and instruction atomicity. |

Access kinds: `MMU_FETCH` only VA `000–0FFh`, `MMU_READ`/`MMU_WRITE` only `100–1DFh`. The I/O area `1E0–1FFh` is not accessible through ordinary fetch/LD/ST, including VA `511`. Device access through `IN/OUT` belongs to P13/P14 and bypasses PTE[15]. Physical frame 7 remains reserved. Errors: `LCA_ERR_MMU_RANGE` for VA outside 0..511, `LCA_ERR_MMU_PROTECTION` for forbidden regions, `LCA_ERR_PAGE_FAULT` for a valid absent page, `LCA_ERR_MMU_MAPPING` for conflicting or reserved mapping, `LCA_ERR_MMU_STATE` for stale mapping/dirty eviction.

## Expected examples / Ejemplos esperados

| VA | Access | VPN/offset | Example mapping | Result |
| --- | --- | --- | --- | --- |
| `0` | FETCH | 0 / 0 | VPN0→PFN0 | PA `0` (resident, different from absent). |
| `31` | FETCH | 0 / 31 | VPN0→PFN0 | PA `31`. |
| `32` | FETCH | 1 / 0 | VPN1→PFN1 | PA `32`. |
| `255` | FETCH | 7 / 31 | VPN7→PFN3 | PA `127`. |
| `256` | READ | 8 / 0 | VPN8→PFN4 | PA `128`. |
| `300` (`12Ch`) | READ/WRITE | 9 / 12 | VPN9→PFN2 | PA `76` (`4Ch`). |
| `511` (`1FFh`) | READ/FETCH | 15 / 31 | reserved | PROTECTION; no physical translation. |
| `-1`, `512` | any | outside 0..511 | — | VA_OUT_OF_RANGE. |

The panel's example maps VPN0→PFN0, VPN1→PFN1, VPN7→PFN3, VPN8→PFN4, VPN9→PFN2. It loads bytes `10 00 2A 00` at physical `20–23h`, sets `PC=20h` and `AX=2Ah`. FETCH reads VA `20–23h` through P05/P06, puts `10h` in IR, sets VPN1.ref, and leaves PC at `20h`. ST writes AX to VA `12Ch`→PA `4Ch`, then sets VPN9.ref/dirty. This sample mapping is chosen for P07 and differs from P06's standalone example `020h→040h`; no PTE is silently assumed.

## Desktop Excel reproduction / Reproducción

1. Start with the newest `develop` workbook. In a **copy** of `workbook/LCA.xlsm`, import `LcaMmu.bas`, `LcaMmuPath.bas`, `LcaMmuPanel.bas`, `LcaMmuTests.bas` through `Alt+F11` → **File/Archivo → Import File/Importar archivo**. Compile with **Debug/Depurar → Compile VBAProject**.
2. `Alt+F8` → `VerificarMmuP07`; record the observed PASS/FAIL and number of assertions. The test resets CPU, buses, RAM and MMU. Run `VerificarRamP06`, `VerificarBusesP05` and `VerificarCpuAlu` as regressions, recording actual results.
3. `Alt+F8` → `ConstruirPanelPaginas`. Confirm 16 rows, VPN15 reserved, absent PFN displayed as `-`, VPN0 present at PFN0. The `300`/`READ` input → **TRADUCIR** should show PA `04Ch` (76) with no dirty/ref side effect. Enter `511`/`READ`: expect a protection error; enter an unmapped valid code VA such as `64`/`FETCH`: expect PAGE_FAULT with PA=-1.
4. Press **EJEMPLO** to restore the mapping, then **FETCH REAL** and **ST REAL**. Inspect the last bus trace and PTE[1] (`ref=1`) / PTE[9] (`dirty=ref=1`). Capture the 16-entry panel and a trace. The test and the panel each reset state for their own demonstration.
5. Save this copy as `workbook/LCA.xlsm` in the repository. Export the four modules back from the saved workbook and compare with `src/vba` including `Attribute VB_Name`. Correct differences and save again. One integrator edits the binary per turn.

The delivered code passed structural checks and `git diff --check`; desktop Excel/VBA execution and workbook synchronization **must be recorded on Windows**. Expected results above are not claimed as observed.

## Review / Revisión

PR `feature/javier/p07-tabla-paginas` → `develop`, `Refs #7`, reviewer Kevin. Include actual compile/test results and screenshots, especially absent VPN versus mapped PFN0, VA300→PA76, and VA511 protection. Keep the issue open until the review and merge.
