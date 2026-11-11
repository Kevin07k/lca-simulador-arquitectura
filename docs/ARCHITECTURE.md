# P02 — Simulator architecture and module contracts (review proposal)

**Status:** Javier's proposal for Kevin's review in issue [#2](https://github.com/Kevin07k/lca-simulador-arquitectura/issues/2). PR approval establishes the team's agreement. Platform: Microsoft Excel with VBA (`.xlsm`). The workbook, VBA modules and UI belong to subsequent tasks.

## 1. Parameters and memory map

| Element | Contract |
| --- | --- |
| Data and data bus | 8-bit unsigned (00–FFh); arithmetic wraps modulo 256. |
| Virtual address and PC | 9 bits (000–1FFh, 512 bytes). `PC` points to the start of an instruction. |
| Physical address and address bus | 8 bits (00–FFh, 256 bytes). |
| General registers | `AX`, `BX`, `CX`, `DX`: 8 bits; codes 0, 1, 2, 3. |
| Control registers | `PC`, `MAR_V`: 9 bits; `MAR_P`, `MDR`, `IR`, `FLAGS`: 8 bits; `IR` holds the opcode. `ZF`, `CF`, `SF` are bits in `FLAGS`. |
| Paging | 32 bytes/page; 16 virtual pages (`VPN` 0–15) and eight physical frames (`PFN` 0–7). `VPN = VA >> 5`, offset `VA & 1Fh`; `PA = PFN*32 + offset`. |

| Virtual | Purpose | Policy |
| --- | --- | --- |
| 000–0FFh (VPN 0–7) | Code: 64 four-byte instructions | Fetch allowed; CPU writes prohibited. Initially populated by loader. |
| 100–1DFh (VPN 8–14) | Data: 224 bytes | CPU reads/writes allowed; fetch prohibited. |
| 1E0–1E7h | Eight-byte keyboard port/buffer | Device access through `IN` only; not addressable by `LD/ST`. |
| 1E8–1EFh | Eight-byte LCD port/buffer | Device access through `OUT` only; not addressable by `LD/ST`. |
| 1F0–1FFh | I/O reserve | Direct CPU access prohibited. |

| Physical | Purpose |
| --- | --- |
| 00–DFh (PFN 0–6) | Seven replaceable frames for VPN 0–14. |
| E0–E7h (PFN 7) | Eight bytes of the circular keyboard queue. |
| E8–EFh (PFN 7) | Eight bytes of the LCD 8×1 video buffer. |
| F0–FFh (PFN 7) | I/O reserve; never allocated to a user page. |

VPN 15 is reserved, has no allocatable present PTE, and is excluded from replacement; its offsets 0–15 correspond to physical frame 7 **only through the I/O controller**. Normal page translation and caches bypass I/O. Virtual regions do not imply contiguous physical memory: any VPN 0–14 may reside in any of the seven free frames.

## 2. ISA and visible state

P04 concretizes the FLAGS layout: bit 0 = ZF, bit 1 = CF, bit 2 = SF; bits 3–7 remain zero. See the [CPU/ALU API and verification](P04_CPU_ALU.md).

**Fixed four-byte instruction**, in increasing address order: `[opcode, R, X, H]`. `R` is register code 0–3 or zero if unused; `X` is an eight-bit immediate, a second register code, or the low address byte; `H` is zero except as the high byte of a virtual address, where only 00h or 01h is valid. Operand address = `256*H + X`. All unused fields must be zero; noncanonical fields cause `INVALID_ENCODING`. The assembler accepts two- or three-digit hexadecimal with an `h` suffix and labels resolved to absolute nine-bit addresses. After the last validated instruction, the code image is filled with `HLT` (`01 00 00 00`); code may not exceed 256 bytes. Default entry point is 000h.

| Opcode | Mnemonic | R | X / H | Effect on completion |
| --- | --- | --- | --- | --- |
| 00h / 01h | `NOP` / `HLT` | 0 | 0 / 0 | NOP advances; HLT stops with PC at the HLT. |
| 10h | `MOV R,#imm` | destination | immediate / 0 | R = immediate. |
| 11h | `MOV R,R2` | destination | source register / 0 | R = R2. |
| 12h | `LD R,[VA]` | destination | low byte / high byte | R = data byte at VA. |
| 13h | `ST [VA],R` | source | low byte / high byte | data byte at VA = R. |
| 20h / 21h / 22h | `ADD R,R2` / `SUB R,R2` / `CMP R,R2` | first register | second register / 0 | ADD/SUB save the result in R; CMP only updates flags from R−R2. |
| 23h / 24h | `INC R` / `DEC R` | register | 0 / 0 | R = R±1. |
| 30h / 31h / 32h / 33h / 34h | `JMP VA` / `JZ VA` / `JNZ VA` / `JC VA` / `JS VA` | 0 | low byte / high byte | PC = target when the condition holds. |
| 40h / 41h | `IN R` / `OUT R` | register | 0 / 0 | Consume keyboard / write LCD; implicit ports. |

Instructions start at addresses divisible by four and always finish within one 32-byte page. Before execution, the decoder checks opcode, fields, register codes and addresses: `LD/ST` only 100–1DFh; jumps only 000–0FCh and four-byte aligned. An unaligned PC or a PC outside code is an error. There is no indirect access or variable-length instruction. The loader validates the entire program before replacing the current image.

`ZF = 1` if the eight-bit arithmetic result is zero; `SF` copies bit 7; `CF` indicates carry on ADD/INC or borrow on SUB/CMP/DEC. All five arithmetic opcodes update all three flags, including CMP, which does not write R. MOV, LD, ST, branches, IN, OUT, NOP and HLT preserve flags. A branch not taken advances four bytes. Normal completion advances `PC += 4`, except a taken jump or HLT. Advancing beyond code causes `PC_OUT_OF_CODE` on the next fetch.

## 3. Translation, replacement, and errors

The table has 16 entries with `VPN, P (present), PFN, dirty, ref`. VPN 0–14 start with `P=0`, `dirty=ref=0`, `PFN` empty; VPN 15 is marked `RESERVED`. Backing storage ("swap") contains exactly 32 bytes for each VPN 0–14. The loader initializes code and data there. For a normal VPN access: validate region/permissions → inspect PTE → if absent, handle an internal `PAGE_FAULT` → translate → inspect cache → access RAM. `PAGE_FAULT` is visible in the trace; then retry the **same** micro-operation without repeating any committed side effect.

Use the lowest free PFN first. If none is free, use CLOCK on PFN 0–6, initial hand 0: for `ref=1`, clear it and advance; the first `ref=0` is the victim, and leave the hand at the next frame. Before reusing a PFN, invalidate all of its lines in L1I, L1D and L2; if `dirty=1`, copy exactly 32 bytes from RAM into the victim VPN's backing storage; invalidate the victim PTE. Load 32 bytes of the incoming VPN from backing storage into the PFN, set `P=1`, PFN, `dirty=0`, `ref=0`, and retry; the completed access sets `ref=1`, and a write sets `dirty=1`. Seven replaceable frames let eight different VPNs trigger replacement.

Fatal errors: `VA_OUT_OF_RANGE` (VA <0 or >511), `PC_OUT_OF_CODE`, `UNALIGNED_PC`, `PROTECTION` (fetch outside code, data outside data or direct I/O access), `INVALID_OPCODE`, `INVALID_ENCODING`, `INVALID_REGISTER`, `INVALID_BRANCH_TARGET`, and `BUS_CONTENTION`. Fatal errors stop RUN/STEP, leave `PC` at the offending instruction, display the code and relevant address, and do not commit that instruction. A valid page fault is recoverable. `IN_EMPTY` is a wait rather than an error: PC/register/flags remain unchanged and RUN pauses until input arrives; STEP may retry. Direct access to I/O addresses is never silently converted to `IN/OUT`.

## 4. Caches and module contracts

The normal path is **CPU → translator → L1I (fetch) or L1D (data) → unified L2 → physical RAM**. Cache tags always use physical addresses. All caches use four-byte lines, direct mapping and replacement by overwrite: L1I = 2 lines (8 B), L1D = 2 lines (8 B), L2 = 8 lines (32 B). `line = PA >> 2`; `index = line mod line_count`; `tag = line div line_count`. No line crosses a page/frame boundary. On an L1 miss, consult L2; on an L2 miss, fill the entire line from RAM. L1I and L1D have separate hit/miss counts; L2 counts one lookup per L1 miss, classified as hit or miss. Counters increment per **CPU-requested byte**, not per line-fill byte.

Data ST operations use *write-through* to L2 and RAM: on miss, fill the line (*write-allocate*), then update L1D, L2 and RAM. RAM is authoritative for swap. Code is immutable during RUN; the loader rebuilds backing storage and invalidates all cache lines before a new program. I/O is uncached; accesses with `IO=1` go directly to E0–EFh. Reset and frame reuse invalidate the relevant lines before reused RAM can be observed.

| Module and logical operation | Inputs → outputs/effects |
| --- | --- |
| CPU `Fetch(PC)` | Valid PC → 4 ordered bytes, IR=first byte, `MAR_V=VA`, `MAR_P=PA`, `MDR=byte` for each read; PC does not advance until instruction completes. |
| Translator `Translate(VA,type)` | Nine-bit VA, `type=FETCH/READ/WRITE` → eight-bit PA or recoverable `PAGE_FAULT` or range/protection error; sets ref on completion, dirty after writing. |
| Cache `Read/Write(PA,type,value)` | Physical PA, `type=FETCH/DATA`; returns byte/status and hit/miss event; never receives VA or I/O. |
| RAM `Read/Write(PA,value)` | PA 00–DFh for pages; E0–EFh only for I/O controller; returns byte/status. F0–FFh forbidden. |
| Bus `Transfer(source,target,PA,value,signals)` | Physical address and byte, one active driver, enabled receiver; records each `READ/WRITE`, `Clock`, `IO`, `Enable/Load` cycle and `High-Z` when undriven; two drivers → `BUS_CONTENTION` before a write. |
| I/O `IN(R)` / `OUT(R)` | `IN` reads FIFO head once and loads R; `OUT` writes R to video and advances cursor; `En_KBD`/`Ld_LCD` signals and events use `IO=1`. |

At each `Clock` edge, the bus records one transfer cycle. Idle: `Read=Write=IO=Enable=Load=0`, no driver, data `High-Z`. RAM read (including a line fill): `Read=1, Write=0, IO=0`, RAM drives, cache/MDR receives; ST write: `Read=0, Write=1, IO=0`, CPU/cache drives, RAM receives. IN: `Read=1, Write=0, IO=1, En_KBD=1`, keyboard drives, register receives; OUT: `Read=0, Write=1, IO=1, Ld_LCD=1`, register drives, LCD/RAM receives. `Enable` selects exactly one driver, and `Load` enables the receiver; `Read` and `Write` are never both 1. A cache hit creates no external RAM cycle and leaves that bus High-Z. Trace fields include microstep, VA if applicable, PA, byte, source, destination and signals; arbitration detects competing drivers before any write.

The keyboard queue is an eight-byte physical circular FIFO at E0–E7h, with `head`, `tail` and `count` (0–8) as controller state. On full, **reject** the new byte and show `KBD_FULL` without changing the queue. Empty `IN` does not advance PC or head; with data, read E0h+head and only upon committing R advance head, decrement count, and advance PC. `En_KBD=1` during the read. LCD has eight bytes E8–EFh initialized to space (20h) and cursor 0–7: `OUT` writes E8h+cursor, sets `Ld_LCD=1`, updates the visible cell and advances cursor modulo 8. Character eight fills the last cell; character nine replaces the first. `OUT` does not modify the queue or flags. Port state and its physical RAM stay coherent.

Reset/validated load: clear general registers and flags; set PC to entry point 000h; deactivate signals and set buses to High-Z; clear queue, pointers and LCD (eight spaces, cursor 0); invalidate caches and zero counters; clear RAM 00–DFh and set PTEs nonpresent; restore the initial program/data backing image (loaded code and zero data). HLT and errors stop RUN; Reset clears them. A recoverable fault preserves PC and avoids a duplicate write on resume.

## 5. Manual review cases

**Translation.** PTE[10] = `(P=1, PFN=3, dirty=0, ref=0)`; VA `14Dh` (333) yields VPN `10`, offset `0Dh` (13), PA `6Dh` (109). This is a valid data read; `ref` becomes 1. VA `1E8h` is reserved for LCD; `LD` must produce PROTECTION, whereas `OUT` writes E8h+cursor without consulting PTE[15]. VA `200h` produces VA_OUT_OF_RANGE.

**Fetch.** PTE[1] = `(P=1, PFN=2)`, PC=`020h`, RAM `40–43h = 10 00 2A 00`. With initially empty caches, four byte fetches translate `020–023h → 040–043h`. First byte: L1I miss, L2 miss, four-byte line fill from RAM; remaining three: L1I hit. Decode `MOV AX,#2Ah`; upon completion, AX=`2Ah`, PC=`024h`, flags unchanged. Derived counts: L1I hit=3/miss=1; L2 hit=0/miss=1. The bus trace has four physical fill reads (`040–043h`) with READ=1, IO=0, and never two drivers. This is a manual expected result, not an executed test of a workbook that does not yet exist.

**I/O.** Initial state: empty queue, eight LCD spaces, cursor=0. Enqueue `A` (`41h`), writing E0h and setting count=1. `IN AX` with `En_KBD=1`, READ=1, IO=1 reads E0h; committing AX=`41h` sets count=0, head=1. `OUT AX` with `Ld_LCD=1`, WRITE=1, IO=1 writes E8h=`41h`; LCD[0]=`A`, cursor=1. Another IN with no key waits without consuming data or advancing PC. These events can be checked when #13 and #14 are implemented.

## 6. Scope and team decision

This document proposes contracts for issues #3–#14; the RAM/cache capacities and eight-byte keyboard/video buffers follow their acceptance criteria. The exact fixed ISA, ranges, reserved PFN, CLOCK algorithm, write-through policy, FIFO/cursor control and error codes are choices proposed here and require Kevin's review before issue #2 can be marked agreed. If "planning document 02" contains different values, adjust this document and its Spanish counterpart **before** approving the PR. No VBA execution or nonexistent screenshot is claimed.

[Versión en español](ARQUITECTURA_ES.md)
