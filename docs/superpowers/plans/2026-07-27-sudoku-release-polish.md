# Sudoku Release Polish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the five game pages' Quick Guide rendering defect, add concise public-release information to the workbook README, and prove that the formula-only game still passes functional and visual regression checks.

**Architecture:** Keep the existing board, dashboard, Coach, hidden engines, named formulas, and palette unchanged. Add one formula-driven regression contract for the Quick Guide defect, remove the obsolete `AM20` cell text that renders underneath each `QuickGuidePanel` shape, and extend the existing README with two style-matched sections using bounded existing rows. Human-facing README prose is verified by exact read-back and screenshots rather than brittle wording tests.

**Tech Stack:** Microsoft Excel 365 Desktop; connected Excel session; ordinary cells, formulas, data validation, conditional formatting, dynamic arrays, named `LAMBDA` functions, and existing Excel shapes. No VBA, macros, Office Scripts, or external runtime behavior.

## Global Constraints

- The first public release formally supports Microsoft Excel 365 Desktop.
- Excel 2024 and Excel for Web remain unverified until tested in clean environments.
- Runtime behavior must remain entirely inside ordinary Excel cells, formulas, data validation, conditional formatting, dynamic arrays, and named `LAMBDA` functions.
- Do not add gameplay, solving techniques, puzzles, difficulty pages, controls, metrics, colors, or icons.
- Preserve the current Excel-green visual system.
- Do not use whole-sheet autofit or broad restyling.
- Keep all five published pages in their clean startup state: no player entries, Coach `Off`, zero conflicts, candidates `AUTO`, and status `IN PROGRESS`.
- Use live Excel control for workbook reads and writes. Use Office.js only for compact scans or shape APIs unavailable through direct tools.
- Because the workbook is binary and must stay open during live edits, make one workbook commit after the complete verified file is saved, rather than committing partial binary states.
- The live tool surface has no save command. After verification, ask the user to save in Microsoft Excel before staging the workbook file.

---

### Task 1: Add a failing Quick Guide regression contract

**Files:**
- Modify: `Sudoku/sudoku.xlsx`, sheet `_Tests`, range `A147:E148`
- Test: `Sudoku/sudoku.xlsx`, sheet `_Tests`, cells `D148:E148`

**Interfaces:**
- Consumes: existing `_Tests!A141:E145` section and row formatting.
- Produces: one formula-driven contract that turns from `FAIL` to `PASS` when Task 2 lands.

- [ ] **Step 1: Reconfirm the connected target and inspect the source test block**

Use `list_document_sessions(surface="excel")`, select the exact `sudoku.xlsx` session, fetch the `read_ranges`, `copy_range_to`, and `write_range` schemas, then read `_Tests!A141:E145`.

Expected: row 141 is a green section header and row 142 is an ordinary `Expected / Actual / Result` test row.

- [ ] **Step 2: Copy the existing five-row test pattern into the new release section**

Use `copy_range_to` on `_Tests`:

```text
sourceRange: A141:E142
destinationRange: A147:E148
```

Expected: the new rows inherit existing test typography, fills, borders, and formulas before their contents are replaced.

- [ ] **Step 3: Write the RED contract**

Patch these cells with `write_range`:

```text
A147 = RELEASE POLISH CONTRACTS

A148 = Legacy Quick Guide cell text is removed
B148 = No cell text can render underneath QuickGuidePanel
C148 = TRUE
D148 = =AND('01 Easy'!$AM$20="",'02 Medium'!$AM$20="",'03 Hard'!$AM$20="",'04 Expert'!$AM$20="",'05 Master'!$AM$20="")
E148 = =IF(D148=C148,"PASS","FAIL")
```

- [ ] **Step 4: Read the new test results and confirm the RED state**

Read `_Tests!A147:E148`.

Expected before implementation:

```text
E148 = FAIL
```

If any test is already `PASS`, inspect its actual referenced cells and correct only an erroneous test assumption; do not weaken the contract.

---

### Task 2: Remove the Quick Guide spill-through defect on all game pages

**Files:**
- Modify: `Sudoku/sudoku.xlsx`, sheets `01 Easy`–`05 Master`, cell `AM20`
- Test: `Sudoku/sudoku.xlsx`, `_Tests!D148:E148`
- Visual test: each game sheet range `B1:BH30`

**Interfaces:**
- Consumes: the existing, identical `QuickGuidePanel` shape on all five published sheets.
- Produces: a clean Quick Guide rendered only by the shape, with no obsolete underlying cell text.

- [ ] **Step 1: Capture the failing visual and structural baseline**

Read images for `B1:BH30` on all five game sheets. Confirm that `QuickGuidePanel` contains the shared text beginning with `01  ENTER`, and confirm that each page also has the obsolete cell text in `AM20`:

```text
Type 1–9 in green-underlined centers. Press Delete to clear an entry. After using Verify, Return Mode to Off.
```

Expected: the underlying `AM20` text visibly spills through or competes with the shape content, and `_Tests!E148` is `FAIL`.

- [ ] **Step 2: Clear only the obsolete cell contents**

For each published sheet, run `clear_range` with:

```text
range: AM20
clearType: contents
```

Do not clear formats, `AK18`, `AK21`, `AK24`, `AK27`, `AK30`, or the `QuickGuidePanel` shape.

- [ ] **Step 3: Verify the structural fix**

Read `AM20` on all five pages and `_Tests!D148:E148`.

Expected:

```text
All five AM20 cells are blank.
D148 = TRUE
E148 = PASS
```

- [ ] **Step 4: Verify every game page visually**

Capture `B1:BH30` for `01 Easy`, `02 Medium`, `03 Hard`, `04 Expert`, and `05 Master`.

Expected on every image:

- no stray characters to the left or right of `02  CANDIDATES`;
- no clipped Quick Guide body text;
- the board, status panel, Coach, and Quick Guide remain aligned;
- the green palette, clue cells, editable centers, candidates, and borders are unchanged.

If text remains clipped inside the shape, stop and inspect `QuickGuidePanel` width, height, font size, and text margins before making a bounded shape-only adjustment. Do not resize the worksheet or board.

---

### Task 3: Add concise public-release information to README

**Files:**
- Modify: `Sudoku/sudoku.xlsx`, sheet `README`, ranges `A10:J11`, `A18:J18`, `D42`, and `A53:J59`
- Content verification: exact read-back of `README!A10:J11`, `README!A18:J18`, `README!A42:J42`, and `README!A53:J59`
- Visual test: `README!A1:J59`

**Interfaces:**
- Consumes: existing header style from `README!A12:J12` and `README!A44:J44`; existing body/list styles from `README!A38:J42`.
- Produces: visible project rationale, version `1.0.0`, build date `2026-07-27`, support boundary, reset instructions, and Verify performance guidance.

- [ ] **Step 1: Read the workbook formatting guidance before layout edits**

Read the complete live-spreadsheet `style_guidelines.md`. Preserve the existing README typography, green section headers, white body background, bounded column widths, and left alignment.

- [ ] **Step 2: Copy existing styles into the two new README sections**

Use `copy_range_to` within `README`:

```text
sourceRange: A12:J12
destinationRange: A10:J10

sourceRange: A49:J49
destinationRange: A11:J11

sourceRange: A44:J44
destinationRange: A53:J53

sourceRange: A38:J42
destinationRange: A54:J58

sourceRange: A42:J42
destinationRange: A59:J59
```

The copies are for style/layout reuse. All copied cell contents listed in the next step must be overwritten.

- [ ] **Step 3: Write the approved README copy**

Patch these cells:

```text
A10 = WHY THIS IS INTERESTING
A11 = A stateful puzzle interface, live candidates, logical coaching, recursive search and automated tests — implemented entirely with ordinary cells and modern Excel formulas.

A18 = Off = normal · Hint = one step · Trace = explanation · Verify = global search and may take several seconds. Return Verify to Off afterward.

D42 = Microsoft Excel 365 Desktop with dynamic arrays and LAMBDA

A53 = RELEASE INFORMATION
A54 = Version
D54 = 1.0.0
A55 = Build date
D55 = =DATE(2026,7,27)
A56 = Supported
D56 = Microsoft Excel 365 Desktop
A57 = Unverified
D57 = Excel 2024 · Excel for Web (unverified)
A58 = Reset
D58 = Delete entries in green-underlined centers, or reopen a clean copy of the workbook.
A59 = Verify performance
D59 = Global verification may take several seconds; Master can take substantially longer on slower computers.
```

Apply `yyyy-mm-dd` number formatting to `D55` without changing adjacent cells.

- [ ] **Step 4: Set bounded row heights for the new content**

Use `resize_range` on README rows only:

```text
10:10 -> 20 points
11:11 -> 30 points
53:53 -> 20 points
54:57 -> 24 points
58:59 -> 30 points
```

Do not autofit the sheet or change column widths.

- [ ] **Step 5: Read the new content and confirm the GREEN state**

Read `README!A10:J11`, `README!A18:J18`, `README!A42:J42`, and `README!A53:J59`.

Expected:

```text
All written labels, values, the typed build date, and the approved instructions match the Step 3 content exactly.
```

- [ ] **Step 6: Verify README visually**

Capture `README!A1:J59`.

Expected:

- `WHY THIS IS INTERESTING` appears between Start Playing and How to Play;
- the rationale fits without clipping;
- `RELEASE INFORMATION` matches the other green section headers;
- version, date, supported, unverified, reset, and performance rows are readable;
- existing links, build status, project status, spacing, and hierarchy remain intact.

If a long line clips, increase only its row height in a bounded increment or shorten that line without removing required meaning.

---

### Task 4: Run full regression and hand off the saved workbook

**Files:**
- Verify: `Sudoku/sudoku.xlsx`, all 13 worksheets and 51 `SDK_*` names
- Modify after user save: Git index and repository history only

**Interfaces:**
- Consumes: completed workbook changes from Tasks 1–3.
- Produces: a verified, saved, versioned Sudoku workbook ready for the Named LAMBDA documentation phase.

- [ ] **Step 1: Run the existing formula regression gate**

Search `_Tests!E:E` for exact `FAIL` and read the summary used by `README!D45`.

Expected:

```text
No FAIL results.
79 PASS results and 18 SKIP/on-demand checks: two obsolete cell-text checks were intentionally converted to screenshot gates, and one new spill-through regression contract was added.
Existing on-demand tests may remain SKIP while Verify is Off.
```

- [ ] **Step 2: Scan for unintended workbook errors**

Search displayed values across the workbook with this regular expression:

```text
#(REF!|VALUE!|NAME\?|CALC!|SPILL!|DIV/0!)
```

Expected: no matches. Literal documentation of `#N/A` in Formula Reference is allowed; evaluated error values are not.

- [ ] **Step 3: Verify the clean startup state on all five pages**

Read the published-sheet cells used by the existing startup contracts and confirm:

```text
Coach mode = Off
Player answer count = 0
Conflict count = 0
Candidates = AUTO
Status begins with IN PROGRESS
```

Also confirm `_Tests!E132:E140` remains `PASS`.

- [ ] **Step 4: Verify the formula-only and dependency boundary**

Use a compact read-only workbook scan to confirm:

- no external workbook links;
- no newly introduced scripts or macros;
- the five hidden engines and `_Tests` remain hidden;
- all 51 `SDK_*` names remain present;
- no formulas in engines or published game areas were changed by the presentation work.

- [ ] **Step 5: Capture final approval images**

Capture:

```text
README!A1:J59
'01 Easy'!B1:BH30
'02 Medium'!B1:BH30
'03 Hard'!B1:BH30
'04 Expert'!B1:BH30
'05 Master'!B1:BH30
'Formula Reference'!A1:F70
```

Expected: no clipping, overlaps, unexpected color drift, visible formula errors, or test-state residue.

- [ ] **Step 6: Ask the user to save the workbook in Microsoft Excel**

The connected tool surface does not advertise Save or Export. Ask the user to click Save in Microsoft Excel after all checks pass. Do not use Computer Use to save or export the workbook.

- [ ] **Step 7: Confirm the saved file and commit the verified binary**

After the user confirms Save, check the file modification time and Git status, ensure `Sudoku/~$sudoku.xlsx` is not staged, then run:

```bash
git add Sudoku/sudoku.xlsx
git commit -m "feat: polish Sudoku workbook for release"
```

Expected: only `Sudoku/sudoku.xlsx` is included in this implementation commit; the Excel lock file remains untracked and excluded.
