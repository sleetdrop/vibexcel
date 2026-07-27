# Sudoku Release Polish Design

## Goal

Prepare the existing formula-only Sudoku workbook for a clean first public release without expanding its gameplay scope or weakening the formula-only constraint.

The first release formally supports Microsoft Excel 365 Desktop. Excel 2024 and Excel for Web remain unverified until they have been tested in clean environments.

## Project Context

`vibexcel` explores the capability boundary between AI agents and modern Excel formulas. Runtime behavior must remain entirely inside ordinary Excel cells, formulas, data validation, conditional formatting, dynamic arrays, and named `LAMBDA` functions.

The Sudoku workbook already contains five playable puzzles, five isolated calculation engines, a formula reference, a hidden regression suite, a formula-only solver, and uniqueness certification. This phase is release polish, not feature development.

## Scope

### Included

- Repair the visible Quick Guide rendering defect and remove clipped or residual text.
- Apply the same robust Quick Guide structure to all five game sheets.
- Confirm that the five game pages share the same visual and interaction system while retaining their independent puzzles and engines.
- Add concise release metadata and operating guidance to the workbook README:
  - version;
  - build date;
  - supported and unverified environments;
  - reset instructions;
  - Verify performance expectations;
  - a short explanation of why the workbook is technically interesting.
- Preserve the existing Excel-green visual language.
- Preserve the clean initial state: no player entries, Coach `Off`, zero conflicts, candidates `AUTO`, and status `IN PROGRESS`.
- Run functional, formula-error, and visual regression checks after changes.

### Excluded

- New solving techniques, puzzles, difficulty pages, controls, metrics, colors, or icons.
- VBA, macros, Office Scripts, add-in-dependent runtime behavior, or external runtimes.
- A claim of verified Excel 2024 or Excel for Web support.
- GitHub publishing, screenshots, GIF creation, Named LAMBDA comment completion, and Hanoi review. Those are later phases with separate plans.

## Design Approach

Use a surgical release-polish pass rather than redesigning the workbook.

The Quick Guide will remain in the right-hand panel. Its four numbered sections will use bounded text areas that fit within the existing panel width and do not depend on text spilling across hidden structural columns. Copy will be shortened only where necessary to render reliably. The game board, dashboard, Coach control, palette, and formulas will remain unchanged unless a direct dependency requires a small matching adjustment.

The README will retain its existing sectioned green-and-white tutorial style. Release information will be added as compact sections or rows rather than turning the README into a second technical reference. Formula Reference remains the detailed in-workbook architecture document.

## Workbook Components

### Public sheets

- `README`: onboarding, release metadata, compatibility, reset, performance, and project rationale.
- `01 Easy` through `05 Master`: playable pages with the shared Quick Guide treatment.
- `Formula Reference`: unchanged in this phase except for fixes required by verified broken references.

### Hidden sheets

- `_Tests`: primary functional regression gate.
- `_Engine` and difficulty-specific engine sheets: unchanged unless a published-sheet change requires a reference-safe adjustment.

## Interaction and Data Flow

Player entries continue to flow from editable center cells into the matching hidden engine. The engine derives live board state, candidates, conflicts, dead ends, Coach output, Verify output, and dashboard values through formulas. This phase changes only presentation copy and layout around that flow.

Reset remains explicit because formulas cannot imperatively clear player-input cells:

1. Delete entries from green-underlined center cells; or
2. reopen a clean copy of the workbook.

Verify remains an explicit persistent mode. Documentation must tell the player to return the mode to `Off` after use and warn that harder boards may take several seconds, with Master potentially taking substantially longer on slower computers.

## Error Handling and Safety

- Make the smallest local change that fixes each visible defect.
- Do not use whole-sheet autofit or broad restyling.
- Inspect merged ranges and hidden structural columns before changing Quick Guide cells.
- Apply changes in small, independently verifiable batches.
- If a live Excel command fails or times out, inspect the affected range before retrying because Excel operations may be non-transactional.
- Never send writes to a stale or ambiguous connected workbook session.
- Do not save test entries or non-`Off` Coach modes in the release state.

## Verification

The phase is complete only when all of the following hold:

- All five game-page screenshots show no clipped, overlapping, or residual Quick Guide text.
- Quick Guide wording is consistent across all five pages.
- Existing game controls and formulas remain functional.
- `_Tests` contains no `FAIL`; expected on-demand tests may remain `SKIP` when Verify is `Off`.
- Workbook search finds no unintended `#REF!`, `#VALUE!`, `#NAME?`, `#CALC!`, `#SPILL!`, or `#DIV/0!` results.
- All five pages start with empty player entries, Coach `Off`, zero conflicts, candidates `AUTO`, and `IN PROGRESS` status.
- README visibly states version, build date, Excel 365 Desktop support, unverified environments, reset behavior, and Verify performance expectations.
- The workbook still contains no VBA, macros, Office Scripts, external runtime dependency, or external workbook link.

## Later Phases

After this phase passes verification:

1. Complete the missing Name Manager comments for the 37 undocumented `SDK_*` names.
2. Build the Sudoku repository package and public documentation.
3. Produce a hero screenshot and optional demonstration GIF.
4. Test a clean downloaded copy on Windows and evaluate Excel 2024 and Excel for Web.
5. Prepare the GitHub release and community-specific launch text.
6. Review and polish the Hanoi workbook and documentation using the same standards.
