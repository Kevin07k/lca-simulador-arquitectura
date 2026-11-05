# P04 — CPU and ALU / CPU y ALU

<details open>
<summary><strong>English</strong></summary>

## Scope and API

`LcaCpu.bas` owns private CPU state and returns a copy through `CpuSnapshot()`. AX/BX/CX/DX use VBA Byte (0–255), codes 0–3. PC and MAR_V use Long with explicit 0–511 validation; MAR_P, MDR and IR are bytes. `CpuReset` clears CPU state, including all flags. This is the CPU part of reset; other modules will reset buses, memory and I/O.

Use `CpuReadRegister(code)`, `CpuWriteRegister(code,value)` and `CpuWriteControl(code,value)` instead of global variables. Invalid writes raise errors before changing state. Control enum: LcaPC, LcaMAR_V, LcaMAR_P, LcaMDR, LcaIR. PC range storage permits 0–511; P12 must enforce executable region and alignment at fetch, including PC=256 after completing the last code instruction.

FLAGS layout is now explicit: bit 0 = ZF (mask 1), bit 1 = CF (mask 2), bit 2 = SF (mask 4); bits 3–7 are zero. `CpuFlagIsSet(mask)` reads a flag. `CpuSetArithmeticFlags(result,carryOrBorrow)` updates all three after validating the result. Plain register/control writes preserve flags.

`AluExecute(operation,destinationCode,sourceCode)` implements opcode values 20h ADD, 21h SUB, 22h CMP, 23h INC, 24h DEC. Unary operations ignore sourceCode; P12 validates unused instruction fields. It returns the eight-bit result. ADD/SUB/INC/DEC write the destination; CMP only changes flags. All inputs are validated before a state change. Intermediate calculations use Long to avoid VBA Byte overflow; results wrap modulo 256. ZF means zero, SF is bit 7, CF means carry for ADD/INC or borrow for SUB/CMP/DEC. No signed-overflow flag is defined.

The ALU does not advance PC, decode instructions, access worksheets, or bypass buses to access memory. The separate [P04 visual adapter](P04_INTERFACE.md), `LcaPanel.bas`, displays live CPU state and provides manual ALU controls. Full fetch/decode/execute, peripheral integration and the remaining simulator panels belong to later issues.

## Run and reproduce

1. Open `workbook/LCA.xlsm` in desktop Excel and open its VBA project with Alt+F11.
2. Compile via Debug → Compile VBAProject.
3. Return to Excel; Alt+F8 → `VerificarCpuAlu` → Run.
4. Expect `P04 Tests!B1 = PASS`, B2 = 197120, B3 = 1379895, and final CPU state RESET. The macro clears the contents of its own report sheet before each run and resets CPU state on success or failure. It does not reset other simulator modules.

To rebuild, import `LcaInicio.bas`, `LcaCpu.bas`, `LcaAlu.bas`, `LcaCpuTests.bas`, `LcaPanel.bas` and `LcaDiagrama.bas` using the [VBA workflow](VBA_WORKFLOW.md). Avoid duplicate module names. Tests and presentation are separate from the engine; only these adapters write to worksheets.

## Verification performed

On November 4, 2026 (America/La_Paz), all three new source files were imported through the VBA editor into the versioned workbook, the compile action completed without an error dialog, and `VerificarCpuAlu` ran in desktop Excel. The observed report was **PASS**, 197120 arithmetic cases and 1379895 assertions. The workbook was saved with that report.

The three modules were exported back through the VBA editor to an ignored temporary folder and compared with the source. VBA normalized identifier casing (for example FLAGS/flags); the source was synchronized to the actual exports, and all three final comparisons matched. The saved workbook's report was also read back from the XLSM package: PASS, 197120 cases, 1379895 assertions.

The tests cover every ordered pair of bytes for ADD/SUB/CMP, all byte inputs for INC/DEC, register aliasing, four register codes, independent result/flag expectations, source/unrelated-register preservation, PC preservation, reset, snapshot isolation, invalid input errors without mutation, and nine-bit values 256/511. They include boundaries 0/127/128/255. Javier must independently reproduce these results and review the PR before integration; this record does not claim his approval.

</details>

<details>
<summary><strong>Español</strong></summary>

## Alcance y API

`LcaCpu.bas` mantiene estado privado y entrega una copia con `CpuSnapshot()`. AX/BX/CX/DX son Byte de VBA (0–255), códigos 0–3. PC y MAR_V son Long validados entre 0 y 511; MAR_P, MDR e IR son bytes. `CpuReset` borra el estado CPU y los flags. Los otros módulos deberán completar el reset de buses, memoria y E/S.

Usar `CpuReadRegister(codigo)`, `CpuWriteRegister(codigo,valor)` y `CpuWriteControl(codigo,valor)` en lugar de variables globales. Las escrituras inválidas generan error antes de modificar el estado. Enumeración de control: LcaPC, LcaMAR_V, LcaMAR_P, LcaMDR, LcaIR. El almacenamiento permite PC de 0–511; P12 valida región ejecutable y alineación al hacer fetch, incluido PC=256 tras completar la última instrucción de código.

Distribución explícita de FLAGS: bit 0 = ZF (máscara 1), bit 1 = CF (máscara 2), bit 2 = SF (máscara 4); bits 3–7 en cero. `CpuFlagIsSet(mascara)` consulta un flag; `CpuSetArithmeticFlags(resultado,acarreoOPrestamo)` actualiza los tres después de validar el resultado. Las escrituras simples de registros/control conservan flags.

`AluExecute(operacion,destino,origen)` implementa opcodes 20h ADD, 21h SUB, 22h CMP, 23h INC y 24h DEC. Las operaciones unarias ignoran origen; P12 valida los campos no usados. Devuelve el resultado de ocho bits. ADD/SUB/INC/DEC escriben destino; CMP solo actualiza flags. Valida entradas antes de modificar estado. Calcula con Long para evitar overflow de Byte en VBA y normaliza módulo 256. ZF indica cero, SF copia bit 7 y CF indica acarreo en ADD/INC o préstamo en SUB/CMP/DEC. No hay flag de overflow con signo.

La ALU no avanza PC, decodifica instrucciones, accede a hojas ni evita buses para leer memoria. El [adaptador visual de P04](P04_INTERFACE.md), `LcaPanel.bas`, muestra estado CPU real y controles manuales de ALU. El ciclo completo, periféricos y demás paneles corresponden a tareas posteriores.

## Ejecutar y reproducir

1. Abrir `workbook/LCA.xlsm` en Excel de escritorio y su proyecto VBA con Alt+F11.
2. Usar Depurar → Compilar VBAProject.
3. Volver a Excel; Alt+F8 → `VerificarCpuAlu` → Ejecutar.
4. Esperado: `P04 Tests!B1 = PASS`, B2 = 197120, B3 = 1379895 y estado final CPU RESET. La macro limpia el contenido de su propia hoja de reporte en cada ejecución y reinicia CPU al terminar o fallar. No reinicia otros módulos del simulador.

Para reconstruir, importar `LcaInicio.bas`, `LcaCpu.bas`, `LcaAlu.bas`, `LcaCpuTests.bas`, `LcaPanel.bas` y `LcaDiagrama.bas` según el [flujo VBA](VBA_WORKFLOW.md), evitando nombres duplicados. Pruebas y presentación están separadas del motor; solo esos adaptadores escriben en hojas.

## Verificación realizada

El 4 de noviembre de 2026 (America/La_Paz) se importaron los tres archivos fuente nuevos desde el editor VBA al libro versionado; la acción de compilar finalizó sin diálogo de error y se ejecutó `VerificarCpuAlu` en Excel de escritorio. Resultado observado: **PASS**, 197120 casos aritméticos y 1379895 comprobaciones. Se guardó el libro con ese reporte.

Se exportaron los tres módulos desde el editor VBA a una carpeta temporal ignorada y se compararon con la fuente. VBA normalizó mayúsculas/minúsculas de identificadores (por ejemplo FLAGS/flags); se sincronizó la fuente con las exportaciones reales y las tres comparaciones finales coincidieron. También se leyó el reporte desde el paquete XLSM guardado: PASS, 197120 casos y 1379895 comprobaciones.

Las pruebas cubren todos los pares ordenados de bytes para ADD/SUB/CMP, todos los bytes para INC/DEC, origen y destino iguales, cuatro códigos de registro, expectativas independientes de resultado/flags, preservación de origen/registros ajenos y PC, reset, aislamiento de la copia, errores sin mutación y valores de nueve bits 256/511. Incluyen límites 0/127/128/255. Javier debe reproducir los resultados y revisar el PR antes de integrar; este registro no afirma su aprobación.

</details>
