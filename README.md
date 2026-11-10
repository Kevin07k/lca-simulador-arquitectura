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
| P11 | Assembly editor, assembler and loader | In review |
| P12–P22 | Full CPU cycle, I/O, programs, delivery | Pending |

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

- P12 full CPU cycle: execute the programs loaded by P11 through the P10 cached memory path.
- P13/P14 keyboard and LCD I/O, outside paging and caches.
- Demo programs, final documentation and delivery in both languages.

This README will be updated as development progresses.
