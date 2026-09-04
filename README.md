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

## Try the Workbooks

Each workbook is the program: download it, open it in Excel, and inspect the formulas in the same file. No installation or add-in is required.

### Formula-Only Sudoku

[![Formula-Only Sudoku in Excel](Sudoku/Preview.png)](Sudoku/README.md)

Play five independent puzzles with live candidates, conflicts, logical coaching, recursive verification, and uniqueness certification.

[Project guide](Sudoku/README.md) · [Direct workbook download](https://raw.githubusercontent.com/sleetdrop/vibexcel/main/Sudoku/sudoku.xlsx)

### Tower of Hanoi

[![Tower of Hanoi in Excel](Hanoi/preview.png)](Hanoi/README.md)

Choose a disk count and step through a complete Tower of Hanoi solution on a cell-based dot-matrix stage.

[Project guide](Hanoi/README.md) · [Direct workbook download](https://raw.githubusercontent.com/sleetdrop/vibexcel/main/Hanoi/hanoi.xlsx)

## Projects

| Project | Description | Workbook |
|---|---|---|
| [Sudoku](Sudoku/README.md) | A five-difficulty Sudoku game with live candidates, logical coaching, recursive verification, uniqueness certification, and regression tests—implemented with formulas only. | [Download](https://raw.githubusercontent.com/sleetdrop/vibexcel/main/Sudoku/sudoku.xlsx) |
| [Hanoi](Hanoi/README.md) | A step-by-step Tower of Hanoi visualization implemented with formulas only. | [Download](https://raw.githubusercontent.com/sleetdrop/vibexcel/main/Hanoi/hanoi.xlsx) |

## Requirements

Projects use modern dynamic-array formulas such as `LET`, `LAMBDA`, `SCAN`, and `REDUCE`. Exact requirements are documented per project.

The current workbooks are verified on Microsoft Excel 365 Desktop for macOS. Excel 365 for Windows, Excel 2024, and Excel for the Web remain unverified.

Files downloaded from GitHub may open read-only or in Protected View. Enable editing to use the interactive controls; the workbooks do not require macros or add-ins.

## Feedback

This repository is an ongoing lab for formula-only interactive software. Bug reports, compatibility results, formula reviews, and ideas for future experiments are welcome in [GitHub Issues](https://github.com/sleetdrop/vibexcel/issues).

## License

[MIT License](LICENSE)
