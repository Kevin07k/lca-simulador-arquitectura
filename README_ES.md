# LCA — Simulador de Arquitectura de Computadoras

**Idioma:** [English](README.md) | Español

<img src="assets/lca-chip.png" alt="Logo LCA de chip azul" width="160">

Proyecto del segundo parcial de Arquitectura de Computadoras, desarrollado por Kevin y Javier.

## Objetivo

Construir un simulador en Microsoft Excel con macros VBA para visualizar el CPU, los buses tri-estado, la jerarquía de caché, la memoria virtual paginada y los periféricos de teclado y LCD 8x1.

## Estado actual

Construido en **Microsoft Excel con VBA**. El [libro](workbook/LCA.xlsm) muestra todas las partes en una sola hoja **LCA** ([mapa de bandas](workbook/README.md)). Ejecutar `VerificarTodo` con Alt+F8 (o `scripts\Test-Workbook.ps1`) para correr todas las pruebas.

| Tarea | Parte | Estado |
| --- | --- | --- |
| P01–P03 | Flujo, contratos, libro | Integrado |
| P04 | Registros del CPU y ALU | Integrado |
| P05–P06 | Buses tri-estado, micropalabra, RAM | Integrado |
| P07–P08 | Tabla de páginas, fallos de página, CLOCK, swap | Integrado |
| P09 | Cachés L1I, L1D y L2 | Integrado |
| P10 | Caché integrada con paginación | Integrado |
| P11 | Editor ASM, ensamblador y cargador | Integrado |
| P12 | Ciclo CPU completo e ISA, vista de máquina para la defensa | Integrado |
| P13 | Teclado y cola FIFO de ocho bytes | Integrado |
| P14 | LCD 8x1 y buffer de video | Integrado |
| P15 | Demos aritméticos con transporte de datos animado | Integrado |
| P16 | Programa de eco teclado → LCD | En revisión |
| P17 | Demo de fallos de página y caché | En revisión |
| P18–P22 | Integración, documentación, entrega | Pendiente |

## Equipo

- Kevin — [@Kevin07k](https://github.com/Kevin07k).
- Javier — [@Javiertorricos](https://github.com/Javiertorricos).
- Equipo: **LCA**.

## Desarrollo

Este repositorio contiene el código fuente, los archivos de ejecución y la documentación del proyecto. El material académico y la planeación local quedan fuera del repositorio.

Se usa `main` para versiones estables, `develop` para integración y una rama `feature/nombre/id-issue-descripcion` por tarea. El compañero revisa cada Pull Request antes de integrar. Assignees identifica únicamente a quien implementa la tarea.

Los módulos VBA se exportan como archivos de texto para revisarlos en Git y el libro se mantiene sincronizado con el código fuente.

- [Índice de documentación: guía, macro de prueba y banda de cada tarea](docs/README.md).
- [Abrir el libro y sincronizar VBA](docs/VBA_WORKFLOW.md) · [scripts de automatización](scripts/README.md).
- [Fuente VBA por capas](src/vba/README.md), [pruebas](tests/README.md), [libro](workbook/README.md), [programas](programs/README.md) e [identidad](docs/IDENTITY.md).

## Gestión de tareas

- [Gestión del proyecto y flujo de revisión (English / Español)](docs/PROJECT_MANAGEMENT.md).

- [Tablero del proyecto](https://github.com/users/Kevin07k/projects/4): **Backlog → In Progress → Done**.
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues): 22 tareas, con 11 asignadas a cada integrante.
- Cada issue incluye secciones desplegables **English** y **Español**.
- Consultar los requisitos enlazados y las relaciones **Blocked by** antes de comenzar o cerrar una integración.

## Próximos pasos

- P18 controles finales de la hoja; P19/P20 documentación técnica; P21 regresión y ensayo de la defensa.
- P22 entrega y verificación de la entrega.

Este README se actualizará conforme avance el desarrollo.
