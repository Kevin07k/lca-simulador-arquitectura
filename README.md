# LCA — Computer Architecture Simulator

**Language:** English | [Español](README_ES.md)

<img src="assets/lca-chip.png" alt="LCA blue chip logo" width="160">

Computer Architecture second midterm project, developed by Kevin and Javier.

## Objective

Build a Microsoft Excel simulator with VBA macros to visualize the CPU, tri-state buses, cache hierarchy, paged virtual memory, and keyboard and 8x1 LCD peripherals.

## Current status

Built in **Microsoft Excel with VBA**. The [workbook](workbook/LCA.xlsm) shows every part on one **LCA** sheet ([band map](workbook/README.md)). Run `VerificarTodo` with Alt+F8 (or `scripts\Test-Workbook.ps1`) to execute every regression.

| Task | Part | State |
| --- | --- | --- |
| P01–P03 | Workflow, contracts, workbook | Integrated |
| P04 | CPU registers and ALU | Integrated |
| P05–P06 | Tri-state buses, control word, RAM | Integrated |
| P07–P08 | Page table, page faults, CLOCK, swap | Integrated |
| P09 | L1I, L1D and L2 caches | Integrated |
| P10 | Caches integrated with paging | Integrated |
| P11 | Assembly editor, assembler and loader | Integrated |
| P12 | Full CPU cycle and ISA, machine view for the defense | Integrated |
| P13 | Keyboard and eight-byte FIFO queue | Integrated |
| P14 | 8x1 LCD and video buffer | Integrated |
| P15 | Arithmetic demos with animated data transport | Integrated |
| P16 | Keyboard-to-LCD echo program | In review |
| P17 | Page fault and cache demo | Integrated |
| P18 | Worksheet controls (LOAD, STEP, RUN, PAUSE, RESET) and synchronized panels | In review |
| P19 | Spanish technical documentation | Integrated |
| P20 | English technical documentation | In review |
| P21–P22 | Regression and defense rehearsal, delivery | Pending |

## Quick start

1. Requirements: Windows and desktop Microsoft Excel (tested with Excel 16).
2. Open [`workbook/LCA.xlsm`](workbook/LCA.xlsm) and click **Enable Content**. If Windows blocks it: right-click the file → Properties → **Unblock**.
3. On the cover click **VERIFICAR TODO** (run all checks): `ALL PASS` must appear.
4. On the machine bar pick a program from the list, click **CARGAR** (LOAD), then use **PASO** (STEP) / **EJECUTAR** (RUN) / **PAUSA** (PAUSE) / **RESET**. The workbook interface is in Spanish.

The [English technical documentation](docs/TECHNICAL_DOCUMENTATION_EN.md) covers setup and permissions, the architecture (with Mermaid diagrams), memory, opcodes and their lengths, the control word, policies, a step-by-step operating guide with demo instructions, three demonstration traces, troubleshooting and the modules by owner. The Spanish version is [here](docs/DOCUMENTACION_TECNICA_ES.md).

## Team

- Kevin — [@Kevin07k](https://github.com/Kevin07k).
- Javier — [@Javiertorricos](https://github.com/Javiertorricos).
- Team: **LCA**.

## Development

This repository contains the project source, execution files, and documentation. Academic reference material and local planning stay outside the repository.

Use `main` for stable versions, `develop` for integration, and one `feature/name/issue-id-description` branch per task. The teammate reviews each Pull Request before integration. Assignees identifies only the person implementing the task.

Export VBA modules as text files for Git reviews and keep the workbook synchronized with the source.

- [Documentation index: every task guide, test macro and sheet band](docs/README.md).
- [Open the workbook and synchronize VBA](docs/VBA_WORKFLOW.md) · [automation scripts](scripts/README.md).
- [VBA source by layer](src/vba/README.md), [tests](tests/README.md), [workbook](workbook/README.md), [programs](programs/README.md) and [identity](docs/IDENTITY.md).

## Task management

- [Project management and review workflow (English / Español)](docs/PROJECT_MANAGEMENT.md).

- [Project board](https://github.com/users/Kevin07k/projects/4): **Backlog → In Progress → Done**.
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues): 22 tasks, with 11 assigned to each teammate.
- Each issue provides expandable **English** and **Español** sections.
- Check linked prerequisites and **Blocked by** relationships before starting or completing integration.

## Next steps

- P21 regression and defense rehearsal.
- P22 release and submission check.

This README will be updated as development progresses.
