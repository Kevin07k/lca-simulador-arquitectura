# P08 — Page faults, CLOCK and swap / Fallos de página, CLOCK y respaldo

Issue [#8](https://github.com/Kevin07k/lca-simulador-arquitectura/issues/8). Depends on P07 (#7) and on the P02 memory contract. The target is the Excel `.xlsm` workbook.

## Modules / Módulos

| File | Purpose / Propósito |
| --- | --- |
| `LcaSwap.bas` | 15 backing images of exactly 32 bytes, VPN 0–14. `SwapReset`, `SwapWriteByte`, `SwapReadByte`; reserved VPN 15 is rejected. The future loader seeds code/data into this image after reset. |
| `LcaPager.bas` | `PagerReset`, `PagerResolve`, `PagerFetchFour`, `PagerReadByte`, `PagerStoreByte`; lowest available PFN then CLOCK 0–6; dirty writeback; 32-byte load; trace and counters. |
| `LcaPagerPanel.bas` | `P08 Swap` worksheet with 16 PTEs, orange last-fault row, backing byte, counters and paging event trace. Ten reproducible steps on eight distinct VPNs. |
| `LcaPagerTests.bas` | `VerificarPagerP08`: seven frames, eighth page, dirty eviction and reload, permissions, frame 7, byte counts and exactly one bus transfer. |

P07's `MmuTryTranslate` and `MmuTranslate` remain low-level translation APIs: an absent page still returns `False` / raises `LCA_ERR_PAGE_FAULT` there. P07's `MmuFetchFour`, `MmuStoreByte`, `VerificarMmuP07` and demonstration are unchanged. **P08 callers must use `PagerFetchFour`, `PagerReadByte` or `PagerStoreByte`** to resolve a fault and retry. Call `CpuReset`, `MemoryPathReset`, `PagerReset` to initialize the P08 demonstration; `PagerReset` resets RAM, PTEs, swap, CLOCK, counters and paging trace, but does not reset the CPU or bus. Do not call `MmuReset` or `RamReset` separately after seeding the P08 image.

## Exact replacement order / Orden de reemplazo

1. Validate VA and access kind with P07; an invalid address or direct I/O access fails before allocating a frame.
2. If absent, record `PAGE_FAULT` with VA/VPN. Pick the lowest free PFN; otherwise CLOCK starts at PFN 0 and scans PFN 0–6. A `ref=1` is cleared and skipped. On a `ref=0`, choose that frame and leave the hand on the next PFN.
3. Before frame reuse, **P09/P10 must invalidate all physical lines in L1I, L1D and L2 for that PFN**. No cache modules exist at P08; the integration point in `PagerHandleFault` is marked for the cache implementer. Do not enable physical caches without implementing this step. For a dirty victim, copy all 32 RAM bytes into its VPN's backing image, clear dirty, then unmap; a clean victim is simply unmapped.
4. Load all 32 backing bytes into the selected physical frame; map the incoming VPN clean and unreferenced. `PagerResolve` retries P07 translation **once**, recording `RESUME`. The successful read/fetch marks `ref`; a successful bus/RAM store marks `dirty/ref`. Paging image moves are frame maintenance and are shown in the paging trace; they are not fabricated P05 CPU bus microsteps.

The four-byte fetch resolves all addresses **before** starting `MemoryFetchStart`; fixed, aligned P02 instructions cannot cross 32-byte pages. A faulting `ST` resolves translation before `MemoryStorePrepare` / `MemoryStoreCommit` and is never replayed after the bus write. `PC` stays unchanged: P12 owns decoding, register/flag/PC commit and final instruction-level atomicity. A fatal bus failure is propagated, not retried as a page fault. The `PagerReadByte` data-read path currently reads RAM directly after translation; P09/P12 can replace that path when implementing their cache/data-bus behavior.

## Reproducible trace / Traza reproducible

Run `ConstruirPanelSwap` (or `PagerDemoReset` on the existing panel), then click **SIGUIENTE** ten times, or **EJECUTAR 10**. The backing image seeds code VPN0 (`10 00 2A 00`) and data VPN8–14 (`value=VPN+offset`); AX=`A5h`.

| Step | Request | Expected result |
| --- | --- | --- |
| 1 | FETCH VA 0 (VPN0) | Fault, load PFN0, IR=`10h`, PC unchanged. |
| 2 | ST VA 256 (VPN8) | Fault, load PFN1, RAM[32]=`A5h`, VPN8 `dirty=ref=1`; one store bus sequence. |
| 3–7 | READ VPN9–13 | Fault/load PFN2–6. Seven frames full; all seven references set. |
| 8 | READ VPN14 | Eighth page; CLOCK clears seven `ref`, evicts clean VPN0 from PFN0, loads VPN14. |
| 9 | FETCH VPN0 | Evicts dirty VPN8 from PFN1; writes **32 bytes** to swap and reloads VPN0. Backing[8,0]=`A5h`. |
| 10 | READ VPN8 | Evicts VPN9 from PFN2; reloads its backed page and returns `A5h`. |

Expected totals after step 10: 10 page faults, 3 evictions, 1 dirty writeback, 320 loaded bytes and 32 written bytes; cursor CLOCK=3. `VerificarPagerP08` also checks protection of VA 511 and 512, absence of VPN15 backing, PFN7 unchanged, no refault on a hit, and absence of repeated bus cycles. These numbers are **expected**, not observed Excel results.

## Desktop integration and review / Integración y revisión

In a **copy** of `workbook/LCA.xlsm` containing P04–P07 modules, import `LcaSwap.bas`, `LcaPager.bas`, `LcaPagerPanel.bas`, `LcaPagerTests.bas` through `Alt+F11` → **Archivo → Importar archivo**. Compile with **Depurar → Compilar VBAProject**. Run `VerificarMmuP07`, `VerificarPagerP08`, then `ConstruirPanelSwap` and the ten-step demo; repeat P06/P05/P04 regressions. Record the actual PASS/FAIL messages and screenshots of the panel/trace, including dirty writeback and recovered `A5h`. Import missing P06 modules (`LcaRam`, `LcaMemoryPath`, `LcaRamTests`) first if the workbook still lacks them, despite their source being merged in Git.

Save the workbook in the repository and export the four P08 modules back from the saved copy. Compare source and exported text including `Attribute VB_Name`; keep one workbook integrator at a time. The delivered source underwent static checks, but Excel compilation, macro execution and binary synchronization require this desktop step. PR `feature/javier/p08-page-fault-swap` → `develop`; link `Refs #8`, include actual evidence and Kevin's review. The #7 source must be present in `develop` before opening a ready-for-review P08 PR.
