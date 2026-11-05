# P04 — Interactive CPU panel / Panel CPU interactivo

**New / Nuevo:** [CPU block diagram / Diagrama de bloques](P04_DIAGRAM.md). Use DIAGRAM / CONTROLS to move between the architectural view and the dashboard. Usar DIAGRAMA / CONTROLES para alternar entre la arquitectura y el panel.

<details open>
<summary><strong>English</strong></summary>

Open `workbook/LCA.xlsm` and select **CPU**. The dark dashboard provides live decimal, hexadecimal and binary register values; highlighted ZF/CF/SF; PC/MAR_V/MAR_P/MDR/IR; ALU controls and a five-entry recent-action trace.

1. In the light **REGISTER** field select AX, BX, CX or DX.
2. Enter a whole decimal byte (0–255) in **VALUE**, confirm with Enter and click **LOAD**. Double-click or press F2 to edit a merged input when pasting text.
3. Select **DESTINATION** and **SOURCE**, then click ADD, SUB or CMP. INC/DEC only use the destination. CMP displays the subtraction result while preserving the register.
4. **RESET** clears CPU registers and flags; it does not reset unimplemented memory or peripheral modules.
5. **RUN TESTS** executes `VerificarCpuAlu`, refreshes the panel and displays PASS/FAIL. Detailed results remain in `P04 Tests`. Tests reset CPU state.

The panel is a manual CPU/ALU laboratory. Instruction fetch, buses, memory, caches and peripherals will be connected in their own issues. Register loading preserves flags. Each button delegates to `LcaCpu`/`LcaAlu`; the engine still has no worksheet dependency. The trace shows action, destination, prior value, arithmetic result and flags. CMP's trace result is not a register write. History is a saved record, not an execution image to restore.

`LcaPanel.bas` contains presentation code and button callbacks. `MostrarPanelCpu` opens/refreshes the existing panel and fits it to the window. Call `PanelActualizar` after changing CPU state through the low-level API. VBA variables do not persist after closing Excel: open with the saved reset state, then use the controls for a new session. If the visual sheet contains stale values after an external macro, run `MostrarPanelCpu` (Alt+F8).

To rebuild the UI, import `LcaPanel.bas` with the CPU, ALU and test modules, then run `ConstruirPanelCpu`. This recreates only the **CPU** sheet's contents/shapes and resets CPU state and panel inputs/history; it preserves `Inicio` and `P04 Tests`. It changes no security settings and does not require a printer. Font, grid and zoom choices apply to the panel view.

## Manual verification

Checked in desktop Excel: build the sheet, load AX=255 (FFh / 11111111), click INC to obtain AX=0 with ZF=1/CF=1/SF=0, DEC to obtain AX=255 with ZF=0/CF=1/SF=1, and CMP against BX=0 to preserve AX=255 and clear CF. ADD/SUB, RESET and the test button were also checked. The final workbook is saved with CPU reset; the existing arithmetic suite must continue reporting 197120 cases and 1379895 assertions.

</details>

<details>
<summary><strong>Español</strong></summary>

Abrir `workbook/LCA.xlsm` y seleccionar **CPU**. El panel oscuro muestra registros en decimal, hexadecimal y binario; ZF/CF/SF resaltados; PC/MAR_V/MAR_P/MDR/IR; controles de ALU e historial de cinco acciones.

1. En el campo claro **REGISTRO**, seleccionar AX, BX, CX o DX.
2. Introducir un byte decimal entero (0–255) en **VALOR**, confirmar con Enter y pulsar **CARGAR**. Para pegar texto en una entrada combinada, editar con doble clic o F2.
3. Seleccionar **DESTINO** y **ORIGEN**; pulsar ADD, SUB o CMP. INC/DEC solo usan destino. CMP muestra el resultado de la resta sin modificar el registro.
4. **RESET** borra registros y flags CPU; no reinicia módulos de memoria o periféricos pendientes.
5. **VERIFICAR** ejecuta `VerificarCpuAlu`, actualiza el panel e informa PASS/FAIL. El detalle permanece en `P04 Tests`. Las pruebas reinician el estado CPU.

Es un laboratorio manual CPU/ALU. Fetch, buses, memoria, cachés y periféricos se conectarán en sus propias issues. Cargar un registro conserva flags. Cada botón delega en `LcaCpu`/`LcaAlu`; el motor sigue sin depender de hojas. El historial muestra acción, destino, valor anterior, resultado aritmético y flags. El resultado de CMP no representa una escritura al registro. El historial es un registro guardado, no una imagen de ejecución para restaurar.

`LcaPanel.bas` contiene presentación y acciones de botones. `MostrarPanelCpu` abre/actualiza el panel y ajusta el zoom a la ventana. Llamar a `PanelActualizar` tras modificar el CPU mediante la API de bajo nivel. Las variables VBA no persisten al cerrar Excel: abrir con el estado guardado en reset y usar los controles para una sesión nueva. Si otra macro deja valores visuales antiguos, ejecutar `MostrarPanelCpu` con Alt+F8.

Para reconstruir, importar `LcaPanel.bas` junto con CPU, ALU y pruebas y ejecutar `ConstruirPanelCpu`. Recrea contenido/formas únicamente de **CPU** y reinicia CPU, entradas e historial; conserva `Inicio` y `P04 Tests`. No cambia seguridad ni requiere impresora. La tipografía, cuadrícula y zoom se aplican a la vista del panel.

## Verificación manual

Comprobado en Excel de escritorio: construir hoja, cargar AX=255 (FFh / 11111111), pulsar INC para AX=0 con ZF=1/CF=1/SF=0, DEC para AX=255 con ZF=0/CF=1/SF=1 y CMP contra BX=0 para conservar AX=255 y borrar CF. También se verificaron ADD/SUB, RESET y el botón de pruebas. El libro final se guarda con CPU reiniciado; la batería aritmética debe seguir informando 197120 casos y 1379895 comprobaciones.

</details>
