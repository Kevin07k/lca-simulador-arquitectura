# Fuente VBA / VBA source

`LcaDiagrama.bas`: esquema vectorial editable con registros, operandos, ALU y flags conectados al motor. Importarlo junto con `LcaPanel.bas`. [Uso y reconstrucción / Usage and rebuild](../../docs/P04_DIAGRAM.md).

English: `LcaDiagrama.bas` provides the editable CPU block diagram with live register/ALU state. Import it together with `LcaPanel.bas`; see the guide above.

`LcaPanel.bas`: panel interactivo de P04; ejecutar `MostrarPanelCpu` para abrirlo o `ConstruirPanelCpu` para reconstruirlo (reinicia CPU y contenido de su hoja). [Uso bilingüe](../../docs/P04_INTERFACE.md).

English: `LcaPanel.bas` is the P04 interactive dashboard. Run `MostrarPanelCpu` to open it or `ConstruirPanelCpu` to rebuild it (resets CPU and its dashboard sheet). See the bilingual guide above.

`LcaInicio.bas` es el módulo inicial de P03. Ejecutar `VerificarInstalacion` escribe `VBA OK - P03` en `Inicio!B13` del libro que contiene el módulo (`ThisWorkbook`). No modifica otros libros.

P04 añade `LcaCpu.bas` (estado privado, reset y accesos validados), `LcaAlu.bas` (ADD/SUB/CMP/INC/DEC) y `LcaCpuTests.bas` (macro `VerificarCpuAlu`). Consultar [API y verificación bilingüe](../../docs/P04_CPU_ALU.md). Los módulos de motor no acceden a hojas; las pruebas escriben su reporte en `P04 Tests`.

Buses, memoria, MMU, caché, cargador, ciclo completo y E/S se añadirán en sus issues según [los contratos de P02](../../docs/ARQUITECTURA_ES.md).

English: `LcaInicio.bas` is the P03 bootstrap module. `VerificarInstalacion` writes `VBA OK - P03` to `Inicio!B13` in its own workbook. Simulation modules belong to later issues. See the [synchronization guide](../../docs/VBA_WORKFLOW.md).

English: P04 adds `LcaCpu.bas` (private state, reset, validated access), `LcaAlu.bas` (ADD/SUB/CMP/INC/DEC) and `LcaCpuTests.bas` (`VerificarCpuAlu`). Engine modules do not access worksheets; tests write to `P04 Tests`. See the bilingual API above. Buses, memory, MMU, caches, loader, full instruction execution and I/O belong to later tasks.

P05 añade `LcaBus.bas`, `LcaControl.bas`, `LcaBusPanel.bas` y `LcaBusTests.bas`. Importar los cuatro, compilar, ejecutar `ConstruirPanelBuses` y `VerificarBusesP05`. El panel muestra ciclos de bus y micropalabras generados por el motor; desde P06 la demostración usa `LcaRam` en vez de la memoria de prueba original. [Contrato y recorrido bilingüe](../../docs/P05_BUS_CONTROL.md). La ejecución completa de instrucciones corresponde a P12.

English: P05 adds bus arbitration, fetch/store transfer steps, a live bus worksheet and a verification macro. Import all four modules and follow the [P05 integration and test guide](../../docs/P05_BUS_CONTROL.md). P06 replaces the original panel fixture with real RAM; caches remain future work.

P06 añade `LcaRam.bas`, `LcaMemoryPath.bas` y `LcaRamTests.bas`; también actualiza `LcaBusPanel.bas` para mostrar lecturas y escrituras de la RAM real. Importar los módulos nuevos, reemplazar el panel existente, compilar y ejecutar `VerificarRamP06`. [Contrato, límites y prueba bilingüe](../../docs/P06_RAM.md).

English: P06 adds 256 bytes in eight 32-byte frames and wires P05 fetch/store cycles to real RAM. Import the three new modules and replace `LcaBusPanel.bas`. See the [P06 test and integration guide](../../docs/P06_RAM.md). P07/P09/P12 still own translation, caches and full instruction execution.

P07 añade `LcaMmu.bas`, `LcaMmuPath.bas`, `LcaMmuPanel.bas` y `LcaMmuTests.bas`: 16 PTE, traducción VA→PA, página 15 reservada, distinción PFN0/ausencia, fallos de página y panel con 16 filas. Importar los cuatro y ejecutar `VerificarMmuP07` y `ConstruirPanelPaginas`. [Contrato y guía bilingüe](../../docs/P07_MMU.md).

English: P07 adds a page table, translation, a bus/RAM adapter, a 16-row worksheet and verification. Import the four modules; run `VerificarMmuP07` and `ConstruirPanelPaginas`. See the [P07 guide](../../docs/P07_MMU.md). Page replacement belongs to P08.

