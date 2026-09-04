# Hanoi Formula Architecture

`hanoi.xlsx` is both the executable program and the canonical formula source. This document is a map for readers who want to inspect the implementation without duplicating every named formula outside Excel.

## Runtime Boundary

All interaction, state derivation, solver logic, validation, and rendering use worksheet formulas and named `LAMBDA` functions. The workbook contains no VBA, macros, Office Scripts, custom functions, external data connections, external workbook links, or add-in dependency.

## Data Flow

```text
Stage controls
    ↓
HN_CLAMP_DISKS / HN_EFFECTIVE_STEP / HN_MAX_STEP
    ↓
Engine solved-state table and active step list
    ↓
lookup and query helpers
    ↓
Stage title, current move, progress, keyframes, and cell-grid renderer
```

The visible input cells remain ordinary editable Excel cells. Downstream formulas consume normalized effective values so pasted text, decimals, and out-of-range numbers cannot break the display.

## Worksheet Responsibilities

### `Stage`

The user-facing sheet contains the disk and step controls, current move, progress, position summary, legend, keyframes, and the dot-matrix display. The stage itself is a grid of small worksheet cells, so its row heights, column widths, alignment, and glyph choices are part of the rendering logic.

### `Guide`

The portable in-workbook guide explains operation, the formula-only boundary, compatibility, worksheet roles, and where to begin inspecting the implementation. It is intentionally concise so the workbook remains useful when distributed without this repository.

### `Engine`

The inspectable calculation sheet contains linked effective controls, the solved-state table, the active move list, named-formula integration points, and compact release checks. Its stable regions are documented directly on the sheet.

## Named-LAMBDA Layers

- Input constraints: `HN_CLAMP_DISKS`, `HN_EFFECTIVE_STEP`, and `HN_MAX_STEP`
- Solver logic: `HN_MOVE_DISK`, `HN_DISK_POS`, `HN_MOVE_ENDPOINT`, and `HN_STEP_LIST`
- Lookup helpers: `HN_MOVE_DISK_AT`, `HN_MOVE_FROM`, `HN_MOVE_TO`, and `HN_POS_ROW`
- Display helpers: `HN_STAGE_CELL`, `HN_STAGE_TITLE`, `HN_MOVE_LABEL`, `HN_PROGRESS_BAR`, and `HN_POSITIONS`

Open Excel's Name Manager for the exact definitions. Keeping those definitions in the workbook avoids maintaining a second executable implementation in Markdown.

## Solver and State Model

For `n` disks, the engine generates all `2^n - 1` legal transitions. Each row records the step, moved disk, source peg, destination peg, and the position of every supported disk. The current step is a query into that deterministic table; the display does not mutate a separate hidden state.

This arrangement makes the workbook easy to audit: change the controls, find the selected row on `Engine`, and trace the presentation formulas back to the same values.

## Release Checks

The Formula Checks block on `Engine` verifies four invariants for the active disk count:

1. The generated sequence has exactly `2^n - 1` steps.
2. Every transition moves exactly one disk.
3. Every move has distinct source and destination pegs.
4. The final state places every active disk on peg C.

The aggregate result reports `PASS` only when all four checks pass. These are formula contracts, not a separate scripted test runtime.

## Compatibility and Inspection

The released workbook is verified on Microsoft Excel 365 Desktop for macOS and requires modern dynamic-array and named-`LAMBDA` support. Excel 365 for Windows, Excel 2024, and Excel for the Web remain unverified.

For a guided review, start on `Guide`, follow the stable regions on `Engine`, and then inspect the `HN_*` names in Name Manager. The repository [README](README.md) provides the quickest download and usage path.
