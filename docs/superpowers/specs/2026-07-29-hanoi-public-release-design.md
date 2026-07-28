# Hanoi Public Release Design

Date: 2026-07-29  
Status: Approved for implementation planning

## Objective

Prepare the Tower of Hanoi workbook for a public GitHub release and later sharing on Hacker News, Reddit, and similar communities. The release should be approachable for casual users, inspectable by formula developers, visually clean, and self-contained without making Hanoi copy Sudoku's documentation structure.

The workbook remains formula-only: no VBA, macros, Office Scripts, custom functions, external runtimes, or external data dependencies.

## Release Principles

- Treat the workbook as both executable software and a portable notebook.
- Keep the first-run experience focused on the visualization.
- Put enough documentation inside the workbook for offline use.
- Keep implementation details discoverable without duplicating every named formula in prose.
- Make the published workbook free of add-in residue, stale references, and unnecessary personal metadata.
- Preserve the geometry of the cell-based dot-matrix stage.

## Workbook Structure

The release workbook will contain three visible sheets in this order:

1. `Stage` — the default active sheet and interactive visualization.
2. `Guide` — concise user and developer documentation.
3. `Engine` — formula source, solved-state table, maintenance contracts, and compact checks.

### Stage

`Stage` remains the landing page and keeps the existing green control/status panels, dark motion stage, legend, and scrolling keyframes.

Changes are limited to release polish and robustness:

- Remove the personal email address.
- Add a small, unobtrusive pointer to `Guide` and a release version label.
- Keep the initial state at three disks and step zero.
- Change the progress readout from filled bins out of ten to a progress bar plus percentage.
- Remove or relabel the static `Display Mode: Full Move` row because it currently resembles an inactive control.
- Preserve the legend and keyframes.
- Show an explicit message when a pasted value is adjusted to a safe effective value.

The public preview will show a non-zero step where disks occupy multiple pegs. The workbook itself will still open at the start pose.

### Guide

`Guide` will be concise and visually consistent with `Stage`. It will contain:

- a short description of the experiment;
- a 30-second Quick Start;
- the formula-only boundary;
- the responsibilities of `Stage`, `Guide`, and `Engine`;
- an architecture overview from controls to solved states to display formulas;
- the named-LAMBDA groups: input constraints, solver logic, lookup/query helpers, and display helpers;
- compatibility and verification status;
- release version and feedback through GitHub Issues;
- directions for developers who want to inspect the formulas more deeply.

It will not reproduce the full text of all `HN_*` formulas. Exact definitions remain available through Excel's Name Manager and their use remains visible on `Engine`.

### Engine

`Engine` remains visible and inspectable. Its stable contracts stay recognizable:

- linked controls;
- solved-state table;
- active step list;
- named-LAMBDA integration points.

Maintenance text will refer to sheets by their actual names rather than the stale `Sheet2` label. Repetitive row-by-row comments will be removed. Only a small number of comments on meaningful contract anchors will remain.

A compact Formula Checks block will verify the active disk count:

- total steps equal `2^n - 1`;
- each transition moves exactly one disk;
- each move has different source and destination pegs;
- the final state places all active disks on peg C;
- an aggregate result reports `PASS` only when every invariant passes.

This check block replaces the need for a separate test sheet.

## Formula and Input Behavior

The user-facing disk count remains 3–6 and the step remains 0 through the current maximum.

`HN_CLAMP_DISKS` will be aligned with the interface contract. Numeric values will be rounded and clamped to 3–6. Text, errors, or blanks that bypass data validation will safely fall back to 3.

The effective-step helper will safely handle values pasted around data validation. Numeric values will be rounded and clamped to 0 through the current maximum; text, errors, or blanks will fall back to 0.

The visible raw controls remain ordinary editable cells with dropdown validation. Formulas consume only their sanitized effective values.

## External Release Files

The public project directory will be:

```text
Hanoi/
├── hanoi.xlsx
├── README.md
└── preview.png
```

The stable filename `hanoi.xlsx` avoids breaking repository download links when the release version changes. Version information belongs in `Guide`, the README, and GitHub release tags.

The README will contain:

- overview and preview;
- Quick Start;
- key features and formula-only constraints;
- workbook map;
- concise formula architecture;
- compatibility and verification status;
- download and inspection guidance;
- GitHub Issues as the feedback route;
- the repository license context.

There will be no separate Hanoi API document or changelog for this release.

## Public Identity and Package Hygiene

The public workbook and documentation may retain the author credit `Yuan Jiang`. They will not expose the personal email currently displayed on `Stage`.

The release package will remove or normalize:

- personal comment-author and document-property metadata that is not needed for attribution;
- ChatGPT and Excel Labs web-extension or custom-function references;
- stale `Sheet2` wording;
- repeated comments that do not add information;
- temporary sheets, comparison artifacts, and lock files.

The final package must contain no VBA project, Office Script, external workbook link, external data connection, or formula dependency on an add-in.

## Visual Regression Strategy

The motion stage is a character grid whose appearance depends on small cell dimensions. Row heights, column widths, font, zoom, character glyphs, and hidden helper columns are therefore part of functional correctness.

Before editing, capture a clean baseline with:

- the ChatGPT pane closed;
- a recorded Excel window size and zoom;
- the entire visible Stage layout in frame;
- the original workbook preserved by Git;
- representative stage states at the start, a middle step, and completion for at least one disk count.

After editing, repeat the screenshots with the same window size, zoom, active cell, visible range, disk count, and step. Compare:

- disk widths and centering;
- peg alignment;
- base continuity;
- row heights and column widths;
- control, status, legend, and keyframe alignment;
- text clipping and wrapping;
- hidden-helper visibility;
- overall page balance.

The preferred baseline is the original Git workbook plus controlled screenshots. A temporary duplicate sheet may be used only if screenshots cannot isolate a layout difference. Any such sheet must be removed before the release workbook is saved.

## Functional Verification

Verification will cover disk counts 3, 4, 5, and 6 at step zero, an early step, a middle step, and the maximum step. It will also test pasted invalid values, including out-of-range numbers, decimals, text, and blanks.

For every test state:

- the current move and disk positions must agree with the solved-state table;
- the stage title and keyframe highlight must use the effective step;
- no displayed or formula cell may contain an Excel error;
- the Formula Checks aggregate must report `PASS`;
- input validation lists and messages must remain usable.

After the workbook is saved and reopened, repeat a representative functional test and the final visual comparison. Inspect the XLSX package to confirm the absence of prohibited runtime and add-in artifacts.

## Success Criteria

The work is complete when:

- a new user can operate the workbook from `Stage` without external instructions;
- an interested developer can understand the architecture from `Guide`, `Engine`, and Name Manager;
- all Formula Checks pass for supported disk counts;
- the Stage grid matches its baseline geometry with no unintended visual regressions;
- the workbook opens cleanly at three disks and step zero;
- the public directory contains only `hanoi.xlsx`, `README.md`, and `preview.png`;
- workbook and package scans find no formula errors, prohibited runtimes, add-in residue, external links, stale sheet references, unnecessary personal metadata, or temporary artifacts.

## Out of Scope

- Automatic animation or timer-driven playback
- VBA, macros, Office Scripts, or add-in-powered behavior
- Support for more than six disks
- A complete named-formula API reference
- A dedicated tests sheet
- A Hanoi-specific changelog
- Publishing or posting to external communities during this implementation cycle
