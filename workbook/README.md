# Libro inicial / Initial workbook

La hoja **CPU Diagrama** muestra la arquitectura con formas editables y flechas. Los botones DIAGRAMA / CONTROLES enlazan las dos vistas. [Guía bilingüe](../docs/P04_DIAGRAM.md).

English: **CPU Diagrama** shows the architecture with editable shapes and arrows. DIAGRAM / CONTROLS link the two views. See the bilingual guide above.

La hoja **CPU** es el panel interactivo de P04: valores DEC/HEX/BIN, flags, controles de carga/ALU/reset e historial. [Instrucciones bilingües](../docs/P04_INTERFACE.md).

English: the **CPU** sheet is the P04 interactive dashboard: DEC/HEX/BIN values, flags, load/ALU/reset controls and recent-action trace. See the bilingual instructions above.

Abrir `LCA.xlsm` en Microsoft Excel de escritorio. Conserva la portada `Inicio` y el módulo inicial `LcaInicio`, junto a los módulos CPU/ALU de P04 descritos abajo. No se debe renombrar un `.xlsx` a `.xlsm`.

P04 incorpora además `LcaCpu`, `LcaAlu` y `LcaCpuTests`. Ejecutar `VerificarCpuAlu` con Alt+F8 para generar el reporte `P04 Tests`: PASS, 197120 casos aritméticos y 1379895 comprobaciones. [API y reproducción](../docs/P04_CPU_ALU.md). El ciclo completo se implementa en P12.

La macro `VerificarInstalacion` cambia `Inicio!B13` a `VBA OK - P03`. Procedimiento y responsables: [VBA_WORKFLOW.md](../docs/VBA_WORKFLOW.md).

English: open `LCA.xlsm` in desktop Microsoft Excel. It contains the `Inicio` sheet and `LcaInicio` module, not the future simulator. Run `VerificarInstalacion` to write `VBA OK - P03` to `Inicio!B13`. See the synchronization guide.

English: P04 also includes `LcaCpu`, `LcaAlu` and `LcaCpuTests`. Run `VerificarCpuAlu` with Alt+F8: `P04 Tests` reports PASS, 197120 arithmetic cases and 1379895 assertions. See the bilingual API above. Full instruction execution belongs to P12.

