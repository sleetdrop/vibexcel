# vibexcel

AI-assisted experiments at the capability boundary between modern Excel formulas and interactive software.

No VBA. No macros. No Office Scripts. No external runtimes.

## About

`vibexcel` explores what AI agents and modern Excel formulas can build together. Each project is a self-contained workbook: executable artifact, interface, formula source, and portable notebook in one file.

The projects in this repository focus on:

- formula-only implementations
- interactive spreadsheet systems
- algorithm visualization
- puzzles and simulations
- modern Excel formula capabilities
- inspectable formula architecture and executable regression contracts

All runtime behavior stays inside Excel without scripting support.

## Projects

| Project | Description | Workbook |
|---|---|---|
| [Sudoku](Sudoku/README.md) | A five-difficulty Sudoku game with live candidates, logical coaching, recursive verification, uniqueness certification, and regression tests—implemented with formulas only. | [Download](Sudoku/sudoku.xlsx) |
| [Hanoi](Hanoi/README.md) | A step-by-step Tower of Hanoi visualization implemented with formulas only. | [Download](Hanoi/hanoi.xlsx) |

## Requirements

Projects use modern dynamic-array formulas such as `LET`, `LAMBDA`, `SCAN`, and `REDUCE`. Exact requirements are documented per project.

Sudoku v1.0.0 is verified on Microsoft Excel 365 Desktop. Excel 2024 and Excel for Web remain unverified for that release.

## License

[MIT License](LICENSE)
