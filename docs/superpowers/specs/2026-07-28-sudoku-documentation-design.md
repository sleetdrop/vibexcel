# Sudoku Documentation and Formula Snapshot Design

**Date:** 2026-07-28

**Status:** Approved direction; written specification pending final review

## Goal

Publish Sudoku as a self-contained, formula-only Excel notebook while also making its purpose, architecture, and Named LAMBDA implementation discoverable and reviewable on GitHub.

## Product Position

The workbook is not merely a binary download accompanied by repository documentation. It is the executable artifact, user interface, technical notebook, formula source, and portable documentation bundle.

Someone who receives only `Sudoku/sudoku.xlsx` must still be able to:

- understand the project and why it is interesting;
- start and reset a puzzle;
- understand compatibility and performance boundaries;
- inspect the calculation architecture and Named LAMBDA interfaces;
- inspect the hidden engines and regression tests if desired.

GitHub documentation complements this experience. It must not replace or weaken the workbook's self-contained documentation.

## Canonical Source and Synchronization Rule

The workbook is the canonical executable artifact.

- User-facing behavior and operational guidance are authoritative in the workbook README.
- Named LAMBDA definitions are authoritative in Excel's Name Manager.
- The workbook Formula Reference is the canonical human-readable API catalog.
- Repository formula documentation is a versioned, read-only snapshot extracted from the released workbook.
- Formula changes are made in the workbook first. The repository snapshot is regenerated afterward; it is never maintained as a second editable implementation.

The repository documentation must state this rule explicitly.

## Documentation Layers

### 1. Workbook README

The existing `README` worksheet remains the portable entry point. It contains:

- the project premise and formula-only constraint;
- Start Playing navigation;
- player instructions and Coach modes;
- workbook structure;
- compatibility requirements;
- build and test status;
- version, build date, reset instructions, and Verify performance guidance.

The release-polish phase already established the required version `1.0.0`, build date `2026-07-27`, official support for Microsoft Excel 365 Desktop, and the `79 PASS / 18 on demand` status. This documentation phase does not rewrite or expand the workbook README unless verification discovers a factual inconsistency.

### 2. Workbook Formula Reference

The existing `Formula Reference` worksheet remains part of the public workbook. It explains the formula system using a six-column catalog:

- function or layer;
- responsibility;
- inputs;
- return value;
- consumers;
- engineering notes.

It covers core state and validation, basic and advanced reasoning, presentation, dead-end detection, recursive search, live Verify integration, bounded uniqueness certification, Coach safety, player interaction, and fallback behavior.

This worksheet is intentionally more detailed than a typical game's embedded help because Sudoku is also an experiment in Excel as an executable notebook. It should not be removed or replaced by a repository-only API document.

### 3. `Sudoku/README.md`

The project README is optimized for someone viewing GitHub before opening Excel. It will contain:

1. a concise project hook;
2. a clean workbook preview;
3. formula-only and AI-assisted project context;
4. feature highlights;
5. quick-start instructions;
6. Coach mode behavior and reset instructions;
7. supported and unverified Excel versions;
8. performance expectations for Verify;
9. a short workbook-structure table;
10. testing and release-status facts;
11. links to the architecture document and Named LAMBDA snapshot;
12. a note that the workbook itself contains the complete portable README and Formula Reference.

The README must be concise enough for Hacker News and Reddit visitors to understand the project without first downloading the workbook. It does not duplicate the workbook's complete Formula Reference table.

### 4. `Sudoku/FORMULA_ARCHITECTURE.md`

This is a hand-written conceptual guide for technical readers. It explains relationships and design decisions that are awkward to express as a row-by-row API catalog.

Required sections:

- formula-only execution boundary;
- workbook data flow from editable centers to hidden engines and back to the visible UI;
- five isolated puzzle engines;
- candidate calculation and conflict detection;
- logical Coach strategy ordering;
- transparent certified-solution fallback;
- on-demand recursive solving and its performance gate;
- bounded uniqueness certification;
- regression-test architecture;
- limitations imposed by formula-only interaction;
- where to inspect the implementation in Excel.

The document may use one compact Mermaid flowchart. It must not reproduce all 51 formulas or restate every row of the workbook Formula Reference.

### 5. `Sudoku/NAMED_FORMULAS.md`

This file is a generated, versioned snapshot of the workbook's 51 `SDK_*` workbook names. Its purpose is to make the binary workbook searchable, linkable, and reviewable on GitHub.

The snapshot contains:

- workbook file name;
- workbook version;
- SHA-256 of the exact `.xlsx` file;
- extraction date;
- explicit generated/read-only warning;
- formula count, expected to be exactly 51;
- one section per `SDK_*` name in stable alphabetical order;
- the stored LAMBDA formula in a fenced `excel` code block;
- the workbook name comment when one exists.

Storage-only prefixes such as `_xlfn.` and `_xlpm.` may be removed in the displayed formula for readability, provided the document states that this is display normalization and does not change the workbook. XML entities must be decoded correctly. No formula may be truncated or manually paraphrased in this snapshot.

The extraction must fail rather than publish an incomplete snapshot when:

- the workbook cannot be opened as a valid ZIP package;
- `xl/workbook.xml` or the defined-name collection is missing;
- the `SDK_*` count is not exactly 51;
- a selected name has an empty formula;
- the workbook SHA-256 cannot be calculated.

The implementation may retain a small deterministic extraction script in `tools/` so future workbook releases can regenerate the snapshot. The script is a release/documentation tool only; it is not a workbook runtime dependency and does not weaken the formula-only claim.

### 6. `Sudoku/Preview.png`

The preview uses the clean startup state of `01 Easy` and shows the board, status dashboard, Coach controls, candidates, and Quick Guide without editor chrome or temporary state.

The image must be legible on GitHub, contain no clipped or overlapping text, and match the committed workbook. It is a documentation asset, not a redesigned mockup.

### 7. Root `README.md`

The repository README will:

- describe `vibexcel` as an exploration of the capability boundary between AI agents and modern Excel formulas;
- keep the formula-only runtime constraints prominent;
- add Sudoku to the Projects table with links to its README and workbook;
- retain Hanoi;
- state Microsoft Excel 365 Desktop as the currently verified target for Sudoku while avoiding unsupported compatibility claims for every project.

## File Layout

```text
Sudoku/
├── sudoku.xlsx
├── Preview.png
├── README.md
├── FORMULA_ARCHITECTURE.md
└── NAMED_FORMULAS.md

tools/
└── extract_named_formulas.mjs
```

The temporary Excel lock file `Sudoku/~$sudoku.xlsx` remains untracked and must never be committed.

## Duplication Policy

Duplication is allowed only when it improves distribution or reviewability:

- Quick-start facts may appear in both workbook and GitHub README.
- Compatibility, reset, and performance warnings must agree exactly across both surfaces.
- Architecture prose belongs in GitHub; the API catalog belongs in the workbook.
- Exact Named LAMBDA definitions are mirrored only through the generated snapshot.
- The repository must not contain a manually maintained second copy of the Formula Reference table.

## Verification

Before committing the documentation release:

1. Recalculate the workbook SHA-256 and confirm it matches `NAMED_FORMULAS.md`.
2. Run the extraction tool twice and confirm byte-for-byte deterministic output.
3. Confirm exactly 51 non-empty `SDK_*` formulas were exported.
4. Check every relative Markdown link and referenced file path.
5. Confirm `Preview.png` exists and visually matches a clean `01 Easy` startup state.
6. Compare README claims against the workbook: version `1.0.0`, Excel 365 Desktop support, Excel 2024/Web unverified, and `79 PASS / 18 on demand`.
7. Confirm the workbook ZIP remains valid and contains no VBA project or external workbook links.
8. Confirm `Sudoku/~$sudoku.xlsx`, `.DS_Store`, and internal planning documents are not part of the public project commit.

## Hanoi Follow-up Principle

Hanoi will be reviewed after the Sudoku documentation release. It will follow the same workbook-first principle without being forced into Sudoku's documentation volume.

- Its workbook should contain enough guidance to remain understandable when distributed alone.
- A concise formula map should be added only if the implementation complexity justifies it.
- GitHub remains the discovery and preview surface.
- A formula snapshot is useful when the binary workbook otherwise prevents meaningful source review.

Consistency across projects means applying the same documentation philosophy, not requiring identical files or page counts.

## Out of Scope

- changing Sudoku gameplay, puzzles, solving strategies, formulas, or visual design;
- changing the workbook filename;
- claiming support for Excel 2024 or Excel for Web;
- publishing to Hacker News or Reddit in this phase;
- rewriting Hanoi before the Sudoku documentation release is complete;
- adding a runtime dependency, macro, Office Script, or external service to the workbook.
