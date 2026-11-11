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

## Pendiente / Pending

P16 (eco teclado → LCD) y P17 (fallos de página y caché).
