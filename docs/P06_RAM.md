# P06 — Framed RAM and initial bus path / RAM por marcos y recorrido inicial

## What changed / Qué se implementó

`LcaRam.bas` stores **256 physical bytes in eight frames of 32 bytes** (`frame = PA \ 32`, `offset = PA Mod 32`). `RamReadPhysical`/`RamWritePhysical` validate `0..255` for loader, paging and diagnostic access. `RamReadUser`/`RamWriteUser` restrict normal CPU data to `00..DFh`. `RamReadIO`/`RamWriteIO` expose only `E0..EFh` to future I/O controllers. `F0..FFh` exists physically but is reserved and cannot be reached by CPU or device adapters. `RamReset` clears RAM and initializes the LCD buffer `E8..EFh` to eight spaces (`20h`). Error codes distinguish out-of-range (`LCA_ERR_RAM_ADDRESS`) from protected regions (`LCA_ERR_RAM_PROTECTION`).

`LcaMemoryPath.bas` connects **actual RAM** to the P05 control layer: a RAM byte drives each external fetch read, then the internal MDR-to-IR/buffer latch runs on a separate microstep. A store first transfers the register into MDR and writes physical RAM **only after** `ControlCommitStore` has clocked and checked the external bus. The existing `LcaBusPanel.bas` now uses this path and the same real RAM for its ten-step demonstration. `LcaRamTests.bas` checks frame boundaries, protection, data, CPU registers, signals, trace and contention.

`LcaRam` proporciona 256 bytes físicos en ocho marcos de 32. El acceso sin restricciones de zona (`RamReadPhysical`/`RamWritePhysical`) es para carga, reemplazo de páginas y diagnóstico; el recorrido normal de CPU pasa por `LcaMemoryPath` y solo accede a `00..DFh`. La zona `E0..EFh` queda para controladores de E/S y `F0..FFh` permanece reservada. El panel de P05 ahora muestra transferencias hechas contra esta RAM real.

## API and limits / Interfaz y límites

| Action / Acción | Call / Llamada | Effect / Efecto |
| --- | --- | --- |
| Reset | `RamReset`; `MemoryPathReset`; `CpuReset` | Clears RAM, bus sequencing and CPU; LCD bytes become `20h`. This demo reset does not set PTE/cache state. |
| Frame calculation | `RamFrameOf(PA)`, `RamOffsetOf(PA)` | Validates `0..255`; e.g. `31→(0,31)`, `32→(1,0)`, `255→(7,31)`. |
| Physical maintenance | `RamReadPhysical(PA)`, `RamWritePhysical(PA,byte)` | All eight frames, including 255; **do not call from CPU instruction execution**. |
| Normal data | `RamReadUser(PA)`, `RamWriteUser(PA,byte)` | `00..DFh` only; standalone setup may seed these bytes. Runtime CPU fetch/store must go through `LcaMemoryPath`. |
| Future devices | `RamReadIO(PA)`, `RamWriteIO(PA,byte)` | `E0..EFh` only, through the device owner. |
| Fetch | `MemoryFetchStart(VA,PA)`, four pairs of `MemoryFetchReadNext` / `MemoryFetchLatchNext` | Eight clocked bus microsteps, `ControlFetched(0..3)` and `IR`. `MemoryFetchFour` runs the same steps without pauses. |
| Store | `MemoryStorePrepare(VA,PA,R)` then `MemoryStoreCommit` | Internal register→MDR, external WRITE→RAM. No RAM mutation at prepare; commit only after arbitration. |

P06 supplies the **sample** mapping `020h→040h` and `14Dh→06Dh` at its call sites; it does not implement the PTE/MMU (P07). It also does not implement L1/L2 caching (P09), instruction decode/PC commit (P12), or real keyboard/LCD device behavior (P13/P14). Once P07/P09 exist, their controller must provide the translated PA and cache hits, and preserve P05's bus cycles. `PC` stays at `020h` throughout this fetch/store demonstration. The tests use a four-byte image `10 00 2A 00`, which is `MOV AX,#2Ah` under P02, without executing that instruction.

## Reproduce in desktop Excel / Reproducir en Excel de escritorio

1. Work on a copy of the latest `workbook/LCA.xlsm`. In VBA editor (`Alt+F11`) import new `LcaRam.bas`, `LcaMemoryPath.bas`, `LcaRamTests.bas`. Replace the existing `LcaBusPanel` module with this version after exporting a backup if it contains local edits. Do not import a duplicate module. **Debug → Compile VBAProject**.
2. `Alt+F8` → `VerificarRamP06`: expected PASS with the assertions recorded by the macro. Raw writes/reads at PA `0`, `31`, `32`, `255` must retain distinct bytes; raw `-1` and `256` must raise range errors; CPU access to frame 7 must raise protection. The macro resets the CPU, bus and RAM; record the **observed** result, not just this expectation.
3. `Alt+F8` → `ConstruirPanelBuses` (or `BusDemoReset` if already constructed). Click **SIGUIENTE / NEXT** eight times. Expect `IR=10h` on the first latch and four RAM reads `040..043h`; after each latch the external bus is High-Z. Click twice more: `AX→MDR`, then `WRITE PA=6Dh`; `RAM[6Dh]` becomes `2Ah` only at the last click. Capture the last trace and the live control word.
4. Click **CONTENCION / CONTENTION**: the trace reports two drivers and `BUS_CONTENTION`, while `RAM[6Dh]` stays `2Ah`. Execute `VerificarBusesP05` and `VerificarCpuAlu` as regressions. Record actual observed results and screenshots.
5. Save the copy as `workbook/LCA.xlsm` in the repository. Export **all modified/new** modules (`LcaBusPanel`, `LcaRam`, `LcaMemoryPath`, `LcaRamTests`) from that saved book, compare them with `src/vba` including `Attribute VB_Name`, and reconcile. Do not merge `.xlsm` as text. One person integrates the workbook at a time.

The source and documentation were checked statically in this delivery. Desktop Excel/VBA compilation, macro execution, screenshot evidence and binary synchronization require the Windows integration above. Until then, report these outcomes as **pending**, not PASS.

## PR evidence / Evidencia para el PR

Base `develop`, branch `feature/javier/p06-ram-marcos`, `Refs #6`. Attach observed results for physical boundaries, invalid addresses, fetch/WRITE signals, and a screenshot of the real-RAM panel after the tenth step. Ask Kevin to review the CPU connection. Keep the issue open until review and integration.
