# P04 — CPU block diagram / Diagrama de bloques del CPU

<details open>
<summary><strong>English</strong></summary>

Open **CPU Diagrama** in `workbook/LCA.xlsm`, or click **DIAGRAM** on the CPU dashboard. This editable vector view shows a processor outline, AX/BX/CX/DX register bank, operand inputs, a classic notched ALU symbol, result, control-register state and flags.

The layout is inspired by the block diagrams in the team's local academic material: *11. 8085.pdf*, pages 19–20, and *17. Buses_y_control.pptx.pdf*, pages 2 and 9. The drawing is original and uses the LCA eight-bit/nine-bit contracts; it is not an implementation of an 8085/8086. The course PDFs remain outside this repository.

## Operate

1. Click **CONTROLS** to load byte values and choose destination/source registers on CPU.
2. Return with **DIAGRAM**. ADD/SUB/CMP/INC/DEC buttons use the same selected registers and engine as the dashboard.
3. After an operation, A/B display the operands consumed by that operation; the ALU names it and RESULT shows its eight-bit output. The register bank shows current state. INC/DEC use constant 1.
4. The green write-back path is shown for arithmetic writes. CMP changes flags only and hides that path; the destination register remains unchanged.
5. RESET clears CPU and the last-operation display. The test button on CPU runs the arithmetic suite and refreshes both views.

Arrows show logical ALU data dependencies, not cycle-accurate bus transfers. The temporary A/B displays are captured operation inputs, not additional architectural CPU registers. Memory, external buses, timing and micro-control belong to later issues. The green path returns to the register bank, with the destination explicitly named.

## Rebuild and verify

Import `LcaDiagrama.bas` alongside `LcaPanel.bas` and the existing CPU/ALU/test modules. Run `ConstruirPanelCpu` to initialize the controls, then `ConstruirDiagramaCpu`. Both macros recreate only their own sheets' content and shapes; the CPU constructor also resets CPU state. `MostrarDiagramaCpu` opens/refreshes the diagram without resetting it. Keep both presentation modules together because they call each other.

Verified in desktop Excel: diagram INC from reset shows A=0, B=1, AX=1, result=01h and the write-back path; CMP against BX=0 preserves AX=1 and hides write-back; RESET restores zero state. Navigation works in both directions. The arithmetic suite still reports PASS, 197120 cases and 1379895 assertions. The workbook is saved in reset state so its initial display matches a new VBA session. Last-operation variables do not persist between Excel sessions.

</details>

<details>
<summary><strong>Español</strong></summary>

Abrir **CPU Diagrama** en `workbook/LCA.xlsm` o pulsar **DIAGRAMA** en el panel CPU. La vista vectorial editable muestra contorno de procesador, banco AX/BX/CX/DX, entradas de operandos, símbolo clásico de ALU con muesca, resultado, estado de registros de control y flags.

El diseño se inspira en los diagramas del material académico local del equipo: *11. 8085.pdf*, páginas 19–20, y *17. Buses_y_control.pptx.pdf*, páginas 2 y 9. El dibujo es original y respeta los contratos LCA de ocho/nueve bits; no implementa un 8085/8086. Los PDF de clase quedan fuera del repositorio.

## Operar

1. Pulsar **CONTROLES** para cargar bytes y elegir registros destino/origen en CPU.
2. Volver con **DIAGRAMA**. Los botones ADD/SUB/CMP/INC/DEC usan la misma selección y motor del panel.
3. Tras una operación, A/B muestran los operandos consumidos; la ALU indica la operación y RESULTADO su salida de ocho bits. El banco muestra el estado actual. INC/DEC usan constante 1.
4. La ruta verde de retorno aparece cuando la operación escribe un registro. CMP solo cambia flags y oculta esa ruta; el registro destino se conserva.
5. RESET borra CPU y la última operación visible. El botón de pruebas de CPU ejecuta la batería y actualiza ambas vistas.

Las flechas indican dependencias lógicas de ALU, no transferencias de bus por ciclo. A/B son capturas de operandos, no registros arquitectónicos nuevos. Memoria, buses externos, temporización y microcontrol corresponden a tareas posteriores. El retorno verde apunta al banco y nombra explícitamente el destino.

## Reconstruir y verificar

Importar `LcaDiagrama.bas` junto con `LcaPanel.bas` y los módulos CPU/ALU/pruebas. Ejecutar `ConstruirPanelCpu` para iniciar controles y luego `ConstruirDiagramaCpu`. Cada constructor recrea contenido y formas solo de su propia hoja; el constructor CPU también reinicia CPU. `MostrarDiagramaCpu` abre/actualiza sin reiniciar. Mantener juntos ambos módulos visuales porque se llaman entre sí.

Verificado en Excel de escritorio: INC desde reset muestra A=0, B=1, AX=1, resultado=01h y retorno de escritura; CMP contra BX=0 conserva AX=1 y oculta el retorno; RESET restaura cero. La navegación funciona en ambos sentidos. La batería conserva PASS, 197120 casos y 1379895 comprobaciones. El libro se guarda en reset para que la vista inicial coincida con una sesión VBA nueva. Las variables de última operación no persisten entre sesiones de Excel.

</details>
