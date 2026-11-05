# LCA — Computer Architecture Simulator

**Language:** English | [Español](README_ES.md)

<img src="assets/lca-chip.png" alt="LCA blue chip logo" width="160">

Computer Architecture second midterm project, developed by Kevin and Javier.

## Objective

Build a Microsoft Excel simulator with VBA macros to visualize the CPU, tri-state buses, cache hierarchy, paged virtual memory, and keyboard and 8x1 LCD peripherals.

## Current status

P03 setup is prepared in **Microsoft Excel with VBA**. The [initial workbook](workbook/LCA.xlsm) contains a cover and an installation-check macro. CPU, memory and I/O simulation belong to later tasks. Kevin selected the blue chip logo; Javier's review is pending.

## Team

- Kevin — [@Kevin07k](https://github.com/Kevin07k).
- Javier — [@Javiertorricos](https://github.com/Javiertorricos).
- Team: **LCA**.

## Development

This repository contains the project source, execution files, and documentation. Academic reference material and local planning stay outside the repository.

Use `main` for stable versions, `develop` for integration, and one `feature/name/issue-id-description` branch per task. The teammate reviews each Pull Request before integration. Assignees identifies only the person implementing the task.

Export VBA modules as text files for Git reviews and keep the workbook synchronized with the source.

- [Open the workbook and synchronize VBA](docs/VBA_WORKFLOW.md).
- [VBA source](src/vba/README.md), [programs](programs/README.md), [tests](tests/README.md) and [identity](docs/IDENTITY.md).
- [P03 verification](docs/P03_VERIFICATION.md).

## Task management

- [Project management and review workflow (English / Español)](docs/PROJECT_MANAGEMENT.md).

- [Project board](https://github.com/users/Kevin07k/projects/4): **Backlog → In Progress → Done**.
- [Issues](https://github.com/Kevin07k/lca-simulador-arquitectura/issues): 22 tasks, with 11 assigned to each teammate.
- Each issue provides expandable **English** and **Español** sections.
- Check linked prerequisites and **Blocked by** relationships before starting or completing integration.

## Next steps

- Complete Javier's identity and P03 peer review.
- Continue with CPU and buses using the P02 contracts integrated into `develop`.
- Complete the technical documentation, operating instructions, and demo traces in both languages.

This README will be updated as development progresses.
