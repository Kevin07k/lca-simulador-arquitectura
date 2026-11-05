# LCA — Simulador de Arquitectura de Computadoras

**Idioma:** [English](README.md) | Español

<img src="assets/lca-chip.png" alt="Logo LCA de chip azul" width="160">

Proyecto del segundo parcial de Arquitectura de Computadoras, desarrollado por Kevin y Javier.

## Objetivo

Construir un simulador en Microsoft Excel con macros VBA para visualizar el CPU, los buses tri-estado, la jerarquía de caché, la memoria virtual paginada y los periféricos de teclado y LCD 8x1.

## Estado actual

Estado CPU y aritmética de P04 implementados en **Microsoft Excel con VBA**. El [libro](workbook/LCA.xlsm) incluye los módulos fuente y una macro reproducible de pruebas CPU/ALU. Ejecutar `VerificarCpuAlu` con Alt+F8: la hoja `P04 Tests` informa PASS para 197120 casos aritméticos y 1379895 comprobaciones. La ejecución completa de instrucciones, memoria, buses y E/S quedan para tareas posteriores. P04 espera revisión de Javier e integración del PR.

## Equipo

- Kevin — [@Kevin07k](https://github.com/Kevin07k).
- Javier — [@Javiertorricos](https://github.com/Javiertorricos).
- Equipo: **LCA**.

## Desarrollo

Este repositorio contiene el código fuente, los archivos de ejecución y la documentación del proyecto. El material académico y la planeación local quedan fuera del repositorio.

Se usa `main` para versiones estables, `develop` para integración y una rama `feature/nombre/id-issue-descripcion` por tarea. El compañero revisa cada Pull Request antes de integrar. Assignees identifica únicamente a quien implementa la tarea.

Los módulos VBA se exportan como archivos de texto para revisarlos en Git y el libro se mantiene sincronizado con el código fuente.

- [Abrir el libro y sincronizar VBA](docs/VBA_WORKFLOW.md).
- [Fuente VBA](src/vba/README.md), [programas](programs/README.md), [pruebas](tests/README.md) e [identidad](docs/IDENTITY.md).
- [Verificación realizada de P03](docs/P03_VERIFICATION.md).
- [API y verificación CPU/ALU de P04 (English / Español)](docs/P04_CPU_ALU.md).
- [Panel CPU interactivo: controles y verificación visual (English / Español)](docs/P04_INTERFACE.md). Abrir la hoja **CPU** para cargar registros y operar la ALU con botones.
- [Diagrama de arquitectura CPU (English / Español)](docs/P04_DIAGRAM.md): abrir **CPU Diagrama** para ver registros, operandos, ALU, flags y retorno de escritura reales.

## Gestión de tareas

- [Gestión del proyecto y flujo de revisión (English / Español)](docs/PROJECT_MANAGEMENT.md).

- [Tablero del proyecto](https://github.com/users/Kevin07k/projects/4): **Backlog → In Progress → Done**.
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues): 22 tareas, con 11 asignadas a cada integrante.
- Cada issue incluye secciones desplegables **English** y **Español**.
- Consultar los requisitos enlazados y las relaciones **Blocked by** antes de comenzar o cerrar una integración.

## Próximos pasos

- Completar la revisión de identidad y de P03 por Javier.
- Continuar con CPU y buses según los contratos de P02 integrados en `develop`.
- Completar la documentación técnica, las instrucciones de uso y las trazas de demostración en ambos idiomas.

Este README se actualizará conforme avance el desarrollo.
