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

P05 añade `LcaBus.bas`, `LcaControl.bas`, `LcaBusPanel.bas` y `LcaBusTests.bas`. Importar los cuatro, compilar, ejecutar `ConstruirPanelBuses` y `VerificarBusesP05`. El panel muestra ciclos de bus y micropalabras generados por el motor; la demostración inyecta bytes de una memoria de prueba. [Contrato y recorrido bilingüe](../../docs/P05_BUS_CONTROL.md). La RAM real corresponde a P06 y la ejecución de instrucciones a P12.

English: P05 adds bus arbitration, fetch/store transfer steps, a live bus worksheet and a verification macro. Import all four modules and follow the [P05 integration and test guide](../../docs/P05_BUS_CONTROL.md). The demo fixture is not the future RAM/cache implementation.

