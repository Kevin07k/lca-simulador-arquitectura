# P05 — Buses and control word / Buses y micropalabra

## Scope / Alcance

P05 adds four exported VBA modules: `LcaBus.bas` (internal and external bus arbitration), `LcaControl.bas` (fetch and ST transfer steps), `LcaBusPanel.bas` (live Excel worksheet), and `LcaBusTests.bas` (reproducible assertions). The existing P04 CPU/ALU modules stay intact. `ControlLastWord` exposes `R/W/C/IO/EN/LD` for each completed microstep. A driver is High-Z when its bus is floating; a completed idle cycle may have zero drivers. Two active drivers cause `BUS_CONTENTION` on `BusClock`, before the caller can commit a RAM write.

P05 incorpora cuatro módulos VBA exportados. El motor de bus arbitra emisores antes de permitir la escritura. La hoja `P05 Buses` presenta los ciclos reales registrados por el motor, no animaciones. El libro que está actualmente en `develop` todavía no contiene estos módulos: hay que importarlos y guardar una nueva versión `.xlsm` desde Excel de escritorio.

## Engine contract / Contrato del motor

| Operation | Sequence / Secuencia | Owner / Responsable |
| --- | --- | --- |
| Fetch from physical RAM / Lectura RAM | `ControlStartFetch(VA,PA)` once; for indexes 0–3: `ControlReadFetchByte(value)` then `ControlLatchFetchByte()` | P06/MMU/cache supplies the value; P05 clocks the buses and updates `MAR_V`, `MAR_P`, `MDR`, `IR`/four-byte buffer. P12 decodes and advances `PC`. |
| Cache hit / Acierto L1I | `ControlReadCacheByte(value)` then `ControlLatchFetchByte()` | No external RAM cycle; external bus is High-Z. This API is for a byte already supplied by a cache. |
| Store / Escritura | `ControlPrepareStore(VA,PA,registerCode)`; `value = ControlCommitStore()`; **then** write `value` to RAM/cache | P06 owns authoritative physical RAM and write-through cache. P05 verifies arbitration before returning the byte. P12 decodes ST and commits `PC`. |
| Reset / Reinicio | `ControlReset` for buses and sequencing; `CpuReset` for CPU | The full reset coordinator belongs to P12. |
| Panel refresh / Actualización | `BusPanelActualizar` after each cycle | Caller invokes it only if the worksheet has been built. The bus engine has no worksheet dependency. |

`VA` is a validated nine-bit address, `PA` an eight-bit address from the MMU. Fetch starts at aligned code addresses `000–0FCh`; RAM PA starts at a four-byte-aligned address within `00–DFh`. ST uses virtual data `100–1DFh`, RAM PA `00–DFh`. P05 checks these transfer ranges; P06/P12 remain responsible for PTEs, protection, actual memory/cache hits, opcode validation, page faults, and instruction atomicity. `ControlFetched(0..3)` returns four ordered bytes. The control layer does not advance PC.

The low-level bus API is `BusBegin(busId, microstep, VA, PA, Read, Write, IO, [EnKBD], [LdLCD])`, then exactly one `BusDrive(busId, source, byte)` and `BusReceive(busId, target)`, then `BusClock(busId)`. `BusClock` validates contention and records the cycle before returning its byte. `BusFloat` makes the other bus High-Z without inserting a fabricated cycle. `BusSnapshot`, `BusTraceCount`, and `BusTraceLine(0..count-1)` expose the current state and the last 128 cycles. Use `BusBegin(...,False,False,False)` plus `BusClock` for an idle cycle with zero drivers. Optional `EnKBD`/`LdLCD` are only legal with the corresponding I/O read/write; P13/P14 implement queue/cursor side effects. The generic `Enable`/`Load` reflect driver/receiver selection.

## Reproduction / Reproducción en Excel de escritorio

1. Work on a copy of `workbook/LCA.xlsm`, open that copy in Microsoft Excel and import the four `.bas` files via VBA editor `Alt+F11` → **File / Archivo → Import File / Importar archivo**. Do not import duplicates with the same `Attribute VB_Name`. Run **Debug / Depurar → Compile VBAProject / Compilar VBAProject**.
2. `Alt+F8` → `ConstruirPanelBuses`. This creates only the `P05 Buses` sheet and resets the CPU/bus demo fixture; it preserves the P04 sheets. The fixture has `PC=020h`, `AX=2Ah`, bytes `10 00 2A 00` at physical `040–043h`, and byte `00h` at `06Dh`. It is a demo source, **not** the P06 RAM module.
3. Press **SIGUIENTE / NEXT** eight times: external READ and internal LATCH alternate for four bytes. The first latch sets `IR=10h`; the four transferred bytes are `10 00 2A 00`. Every external READ has one RAM driver. During an internal LATCH the external bus reads `High-Z`. `PC` stays `020h` pending P12 decode/commit.
4. Press NEXT twice more: internal `AX -> MDR` then external `CPU/CACHE -> RAM` with `W=1, IO=0`. The fixture at `06Dh` becomes `2Ah` only after the latter edge. `VA=14Dh` maps to `PA=6Dh` in this *configured fixture*, matching P02's sample PTE.
5. Press **CONTENCION / CONTENTION**. The trace records `CPU + OTHER_OUTPUT`, `drivers=2`, `BUS_CONTENTION`; the fixture is unchanged. Press RESET to start again.
6. `Alt+F8` → `VerificarBusesP05` checks all four fetch bytes and CPU control registers, High-Z, bus words, write ordering, trace and contention. The macro resets CPU/bus state and reports a PASS count or the first failure; it does **not** execute P06 RAM/cache, P12 instruction decode, or Excel UI clicks. Restore the panel demo with `BusDemoReset` after running it. Also rerun `VerificarCpuAlu` to check P04 regression.
7. Save the copy as `.xlsm`. Export the four modules back out of that same workbook, compare each with `src/vba` (including `Attribute VB_Name`), reconcile any differences, and commit synchronized source and workbook. A screenshot of the worksheet after step 4 and after contention can accompany the PR.

**Expected, not observed here:** the values and trace above follow the code and P02 contract. A Linux checkout cannot execute desktop Excel/VBA or regenerate the workbook; the integrator must record observed compile, macro, and panel results before declaring the issue accepted.

## PR checklist / Lista para el PR

- `Refs #5`; base `develop`; head `feature/javier/p05-buses-control`.
- Attach the observed `VerificarBusesP05` result, `VerificarCpuAlu` regression result, and panel screenshots after fetch/ST and contention.
- State who integrated the `.xlsm` and confirm the exported modules match the saved workbook.
- Request Kevin's review. Keep #5 in progress until the reviewed PR is integrated.
