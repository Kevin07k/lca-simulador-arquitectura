# Project management / Gestión del proyecto

<details open>
<summary><strong>English</strong></summary>

## Links and responsibilities

- [Repository](https://github.com/Kevin07k/lca-simulador-arquitectura)
- [Project board](https://github.com/users/Kevin07k/projects/4)
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues)
- Kevin (@Kevin07k): P01, P03, P04, P09, P10, P11, P12, P15, P17, P19, P21.
- Javier (@Javiertorricos): P02, P05, P06, P07, P08, P13, P14, P16, P18, P20, P22.
- Assign each issue to its implementer; request the teammate's review on the PR.

## Workflow

1. Check the issue's acceptance criteria and dependencies. Confirm prerequisites are integrated into `develop`.
2. Move the card from **Backlog** to **In Progress** when work starts.
3. Create `feature/<student>/pXX-description` from updated `develop`. There is no permanent branch per student.
4. Commit coherent, verified changes and open a PR targeting `develop`, referencing the issue with `Refs #N`.
5. The teammate reviews the changed files and relevant verification results. Keep the card **In Progress** during review.
6. After approval, integrate using **Create a merge commit** to preserve individual commits.
7. Verify integration and acceptance, close the issue, and move the card to **Done**. A PR into `develop` may require manual issue closure because `main` is the default branch.

`main` contains stable deliveries; `develop` integrates reviewed tasks. Do not push task changes directly to either branch. Create each new feature from updated `develop`, independently of earlier feature branches.

## Dates and verification

Each task needs an implementer and a target-date field in the Project. Actual milestone dates and the official submission deadline are pending confirmation; do not invent them. Before completing P01, manually verify Javier's repository access and the Project's date field and visible statuses: **Backlog → In Progress → Done**.

Verification should fit the change: documentation needs content, link and formatting review; code needs reproducible checks and expected/observed results. Record results in the PR. Screenshots are useful only when they demonstrate a visual configuration or behavior.

## Excel/VBA

Version exported `.bas`, `.cls`, `.frm` and associated `.frx` resources alongside the executable workbook. Coordinate one workbook integrator at a time and keep source and workbook aligned. Never commit Excel lock files such as `~$simulador.xlsm`. Workbook preparation belongs to P03; P01 does not implement the simulator.

</details>

<details>
<summary><strong>Español</strong></summary>

## Enlaces y responsabilidades

- [Repositorio](https://github.com/Kevin07k/lca-simulador-arquitectura)
- [Tablero del proyecto](https://github.com/users/Kevin07k/projects/4)
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues)
- Kevin (@Kevin07k): P01, P03, P04, P09, P10, P11, P12, P15, P17, P19, P21.
- Javier (@Javiertorricos): P02, P05, P06, P07, P08, P13, P14, P16, P18, P20, P22.
- Asignar cada issue a quien la implementa; solicitar la revisión del compañero en el PR.

## Flujo de trabajo

1. Revisar la aceptación y las dependencias de la issue. Confirmar que los requisitos previos estén integrados en `develop`.
2. Pasar la tarjeta de **Backlog** a **In Progress** al comenzar.
3. Crear `feature/<estudiante>/pXX-descripcion` desde `develop` actualizado. No hay rama permanente por estudiante.
4. Guardar cambios coherentes y verificados en commits y abrir un PR hacia `develop`, enlazando la issue con `Refs #N`.
5. El compañero revisa los archivos modificados y los resultados de verificación pertinentes. Mantener **In Progress** durante la revisión.
6. Tras la aprobación, integrar con **Create a merge commit** para conservar los commits individuales.
7. Comprobar integración y aceptación, cerrar la issue y mover la tarjeta a **Done**. Un PR hacia `develop` puede necesitar cierre manual de la issue porque la rama predeterminada es `main`.

`main` contiene entregas estables; `develop` integra tareas revisadas. No subir cambios de tareas directamente a esas ramas. Cada feature nueva nace de `develop` actualizado, independientemente de las features anteriores.

## Fechas y verificación

Cada tarea necesita responsable y un campo de fecha prevista en el Project. Las fechas reales de hitos y el cierre oficial están pendientes de confirmación; no inventarlas. Antes de completar P01, verificar manualmente el acceso de Javier y el campo de fecha y los estados visibles del Project: **Backlog → In Progress → Done**.

La verificación depende del cambio: documentación requiere revisar contenido, enlaces y formato; código requiere comprobaciones reproducibles con resultados esperados y observados. Registrar resultados en el PR. Las capturas sirven cuando demuestran una configuración o comportamiento visual.

## Excel/VBA

Versionar los módulos exportados `.bas`, `.cls`, `.frm` y recursos `.frx` asociados junto al libro ejecutable. Coordinar un integrador del libro por turno y mantener fuente y libro alineados. Nunca incluir temporales de Excel como `~$simulador.xlsm`. Preparar el libro corresponde a P03; P01 no implementa el simulador.

</details>
