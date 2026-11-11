# Programas de demostración / Demo programs

## `arith/` — P15, demos aritméticos / arithmetic demos

Cinco programas en el ensamblador de P11 que usan AX, BX, CX y DX con bucles y terminan en `HLT`. Cada uno guarda su resultado en la memoria de datos (`[100h]`, `[101h]`). Se cargan con los botones del bloque **P15** de la hoja LCA (o con el editor P11) y se ejecutan con PASO o EJECUTAR.

Five programs in the P11 assembly that use AX, BX, CX and DX with loops and end in `HLT`; each stores its result in data memory. Load them with the **P15** block buttons (or the P11 editor) and run with STEP or RUN.

| Archivo / File | Programa / Program | Entrada / Input | Resultado esperado / Expected |
| --- | --- | --- | --- |
| [`01_suma.asm`](arith/01_suma.asm) | SUMA | CX = 5 | AX = 15, `[100h]` = 15 |
| [`02_multiplicacion.asm`](arith/02_multiplicacion.asm) | MULTIPLICAR | BX = 6, CX = 7 | AX = 42, `[100h]` = 42 |
| [`03_fibonacci.asm`](arith/03_fibonacci.asm) | FIBONACCI | CX = 10 | AX = 55, BX = 89, DX = 89 |
| [`04_division.asm`](arith/04_division.asm) | DIVIDIR | AX = 47, BX = 5 | CX = 9, AX = 2 |
| [`05_mcd.asm`](arith/05_mcd.asm) | MCD | AX = 48, BX = 18 | AX = 6 |

La fuente de verdad es `src/vba/LcaArith.bas`: `scripts/Test-Workbook.ps1` comprueba que estos archivos coinciden. Trazas, pasos y banderas: [P15_ARITHMETIC_DEMO.md](../docs/P15_ARITHMETIC_DEMO.md).

The source of truth is `src/vba/LcaArith.bas`; `scripts/Test-Workbook.ps1` checks that these files match.

## `echo/` — P16, eco teclado → LCD / keyboard-to-LCD echo

[`eco.asm`](echo/eco.asm): `IN AX` toma la tecla más antigua de la cola (E0–E7h) y `OUT AX` la escribe en la celda del cursor del LCD (E8–EFh). La tecla `.` termina el programa: no se muestra, `[100h]` = caracteres mostrados y `HLT`. Con la cola vacía la CPU espera en `IN` (WAIT_IN) y continúa al teclear más. Se carga con el bloque **P16** de la hoja LCA.

[`eco.asm`](echo/eco.asm): `IN AX` takes the oldest key from the queue and `OUT AX` writes it to the LCD cell under the cursor. The `.` key ends the program: it is not shown, `[100h]` holds the number of characters shown, and the CPU halts. An empty queue makes the CPU wait on `IN`; it continues when more keys are typed.

La fuente de verdad es `src/vba/LcaEcho.bas` (comprobado por `scripts/Test-Workbook.ps1`). Comportamiento y ejercicio con ocho caracteres: [P16_ECHO_DEMO.md](../docs/P16_ECHO_DEMO.md).

## `memoria/` — P17, faults de página y caché / page faults and cache

[`memoria.asm`](memoria/memoria.asm): siete `ST` a páginas dispersas (faults + misses), un bucle que lee `[100h]` cuatro veces (hits), 10 páginas para 7 marcos (la CLOCK desaloja y guarda las páginas dirty) y la recuperación del dato modificado (`BX = A1h`). Contadores esperados: 30 instrucciones, 11 faltas, 4 evicciones, 3 write-backs, L1D 3/10, L1I 93/27, L2 6/31.

[`memoria.asm`](memoria/memoria.asm): seven scattered `ST` (faults + misses), a loop that reads `[100h]` four times (hits), 10 pages for 7 frames (the CLOCK evicts and writes back dirty pages) and recovery of the modified byte.

La fuente de verdad es `src/vba/LcaMemDemo.bas` (comprobado por `scripts/Test-Workbook.ps1`). Derivación de los contadores: [P17_MEMORY_DEMO.md](../docs/P17_MEMORY_DEMO.md).

## Pendiente / Pending

P17 (fallos de página y caché).
