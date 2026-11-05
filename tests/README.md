# Verificación / Verification

Las evidencias iniciales están en [P03_VERIFICATION.md](../docs/P03_VERIFICATION.md). Para CPU/ALU, ejecutar `VerificarCpuAlu` (Alt+F8) en `workbook/LCA.xlsm`: esperado PASS, 197120 casos y 1379895 comprobaciones en `P04 Tests`. Fuente: `src/vba/LcaCpuTests.bas`. [Pasos y resultados de P04](../docs/P04_CPU_ALU.md). La macro borra el contenido de su propia hoja de reporte y reinicia el estado CPU; no ejecutar durante una simulación que se quiera conservar.

English: initial verification is recorded in [P03_VERIFICATION.md](../docs/P03_VERIFICATION.md). Run `VerificarCpuAlu` (Alt+F8) in `workbook/LCA.xlsm`: expect PASS, 197120 cases and 1379895 assertions in `P04 Tests`. Source: `src/vba/LcaCpuTests.bas`. See the P04 guide above. This macro clears its own report sheet and resets CPU state; do not run during a simulation you need to preserve. Memory and I/O tests belong to their future implementations.

