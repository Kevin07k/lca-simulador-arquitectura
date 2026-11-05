# Fuente VBA / VBA source

`LcaInicio.bas` es el módulo inicial de P03. Ejecutar `VerificarInstalacion` escribe `VBA OK - P03` en `Inicio!B13` del libro que contiene el módulo (`ThisWorkbook`). No modifica otros libros.

Los módulos CPU/ALU, buses, memoria, MMU, caché, cargador y E/S se añadirán en sus issues según [los contratos de P02](../../docs/ARQUITECTURA_ES.md). P03 no implementa esos componentes.

English: `LcaInicio.bas` is the P03 bootstrap module. `VerificarInstalacion` writes `VBA OK - P03` to `Inicio!B13` in its own workbook. Simulation modules belong to later issues. See the [synchronization guide](../../docs/VBA_WORKFLOW.md).

