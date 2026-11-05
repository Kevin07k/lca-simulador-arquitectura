# Libro y fuente VBA / Workbook and VBA source

## Español

**Integrador del libro para P03: Kevin (@Kevin07k). Revisor: Javier (@Javiertorricos).** Kevin incorpora la fuente revisada al `.xlsm`; Javier importa y verifica en una copia. Un solo integrador modifica el binario por turno. Registrar cualquier cambio de turno en el PR antes de modificar el libro.

### Estructura

- `src/vba/`: módulos exportados `.bas`, clases `.cls`, formularios `.frm` y sus `.frx` asociados.
- `workbook/LCA.xlsm`: libro real de Excel, mantenido junto a la fuente del mismo commit.
- `programs/`: programas y datos de demostración, a partir de P15–P17.
- `docs/`: arquitectura, identidad, procedimiento y evidencia.
- `assets/`: logo editable.
- `tests/`: pruebas y recorridos de verificación.

### Abrir y verificar

1. Abrir `workbook/LCA.xlsm` en Excel de escritorio. Usar las macros solo si se confía en este código revisado. No requiere habilitar acceso programático al proyecto VBA.
2. Abrir el editor con `Alt+F11` y seleccionar **VBAProject (LCA.xlsm)**.
3. En **Depurar → Compilar VBAProject**, comprobar que no hay errores.
4. Volver a Excel con `Alt+F11`; `Alt+F8` → `VerificarInstalacion` → Ejecutar.
5. Esperado: `Inicio!B13` muestra `VBA OK - P03`. La macro solo comprueba la instalación; no ejecuta CPU, memoria ni E/S.

### Importar y sincronizar

1. Trabajar en una copia del libro y guardar antes de modificarlo. Comprobar que se seleccionó el proyecto de esa copia en el editor.
2. Si ya existe `LcaInicio`, quitarlo desde el explorador de proyectos. Cuando Excel pregunte si se desea exportar, guardar una copia fuera de `src/vba` si hay cambios que conservar. No importar dos módulos con el mismo nombre.
3. **Archivo → Importar archivo** (`Ctrl+M` en el editor); seleccionar `src/vba/LcaInicio.bas`. Para otros módulos repetir con `.bas`, `.cls` y `.frm`; mantener cada `.frx` junto a su `.frm` antes de importar.
4. El código de `ThisWorkbook` y hojas no se importa como clase nueva: copiar su código en el componente de documento correspondiente y conservar sus nombres de código. P03 no añade código a esos componentes.
5. Compilar, ejecutar la verificación y guardar como **Libro de Excel habilitado para macros (*.xlsm)**, nunca como `.xlsx`.
6. Exportar cada módulo desde el editor (**Archivo → Exportar archivo**, `Ctrl+E`) a una carpeta temporal. Comparar con `src/vba`, incluyendo los atributos `Attribute VB_Name`. Los `.bas` de P03 usan solo ASCII, compatible con la exportación de Excel. No modificar acentos/encoding de otros módulos sin revisar su importación.
7. Si los módulos exportados difieren, revisar el cambio y actualizar la fuente. Guardar el libro desde esa misma versión del código. Adjuntar en el PR el recorrido y resultado real, y subir fuente y libro juntos. No intentar resolver conflictos del `.xlsm` con un merge de texto.

### Reconstruir desde la fuente

Si no se dispone del libro, crear un libro vacío en Excel, dejar una hoja con nombre **Inicio** e importar `LcaInicio.bas` siguiendo los pasos anteriores. Guardar como `.xlsm` y ejecutar la macro: funciona con esa hoja vacía. Para recuperar además la portada diseñada, partir del `workbook/LCA.xlsm` versionado y sincronizar sus módulos. El diseño del libro se conserva en el binario; el comportamiento VBA se conserva en la fuente.

### Revisión cruzada

Javier debe repetir la importación en una copia, compilar, ejecutar la macro y exportar para comparar con la fuente. Registrar su resultado en el PR. Esta revisión y el acuerdo de identidad son condiciones pendientes antes de cerrar #3 y moverla a Done.

## English

### Panel de P04 / P04 dashboard

Importar también `LcaDiagrama.bas` junto a `LcaPanel.bas`; ejecutar `ConstruirDiagramaCpu` después del constructor del panel para recuperar la segunda vista. Also import `LcaDiagrama.bas` with `LcaPanel.bas`; run `ConstruirDiagramaCpu` after the dashboard constructor to rebuild the second view. [Guía / Guide](P04_DIAGRAM.md).

Importar también `LcaCpu.bas`, `LcaAlu.bas`, `LcaCpuTests.bas` y `LcaPanel.bas`. `MostrarPanelCpu` abre/actualiza el panel; `ConstruirPanelCpu` lo reconstruye y reinicia CPU, entradas e historial. Para uso normal, abrir la hoja CPU y usar sus botones. [Instrucciones bilingües](P04_INTERFACE.md).

Also import `LcaCpu.bas`, `LcaAlu.bas`, `LcaCpuTests.bas` and `LcaPanel.bas`. `MostrarPanelCpu` opens/refreshes the dashboard; `ConstruirPanelCpu` rebuilds it and resets CPU, inputs and history. For normal use, open CPU and use its buttons. See the bilingual instructions above.

**P03 workbook integrator: Kevin (@Kevin07k). Reviewer: Javier (@Javiertorricos).** Only one integrator edits the binary at a time; record handovers in the PR.

Open `workbook/LCA.xlsm` in desktop Excel. In the VBA editor (`Alt+F11`), select its project and compile. Run `VerificarInstalacion` via `Alt+F8`: `Inicio!B13` must become `VBA OK - P03`. This only verifies bootstrap installation.

To synchronize, work in a saved copy, remove the old standard module after preserving any changes, then use **File → Import File** (`Ctrl+M`) for `src/vba/LcaInicio.bas`. Import `.bas`, `.cls`, `.frm` as needed and keep `.frx` companions beside `.frm`. Document modules (`ThisWorkbook` and sheets) require code in the matching document component, not a new imported class. P03 has no document-module code.

Compile, run and save as a genuine macro-enabled workbook. Export modules (`Ctrl+E`) to a temporary folder and compare them with `src/vba`, including name attributes. P03 source is ASCII. Review differences and commit matching source and workbook together. Never merge `.xlsm` as text. No programmatic VBA-project trust setting is required for manual import.

To rebuild behavior from source, create an empty Excel workbook with a sheet named `Inicio`, import `LcaInicio.bas`, save as `.xlsm` and run the macro. Use the versioned workbook as the template to retain its cover design. Javier must independently import, compile, run and compare the export, and both teammates must approve the identity before closing #3.
