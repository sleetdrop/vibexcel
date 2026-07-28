# Hanoi Public Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a clean, documented, visually verified, formula-only public release of the Tower of Hanoi workbook.

**Architecture:** Keep `Stage` as the interactive landing page, add a concise `Guide`, and retain `Engine` as the inspectable formula implementation with compact invariant checks. Make workbook edits through the connected Excel session, verify the cell-grid visually before and after, then sanitize and inspect the saved XLSX package before renaming it to the stable public filename.

**Tech Stack:** Microsoft Excel 365 Desktop, modern Excel formulas and named `LAMBDA`s, Codex Excel document control, Computer Use for Excel UI verification, ZIP/XML inspection for XLSX package hygiene, Markdown, Git.

## Global Constraints

- Runtime behavior must use Excel formulas only.
- Do not add VBA, macros, Office Scripts, custom functions, external runtimes, external data connections, or external workbook links.
- Supported disk counts are exactly 3, 4, 5, and 6.
- The workbook must open on `Stage` at three disks and step zero.
- Preserve the cell-grid stage geometry: font, row heights, column widths, character alignment, and hidden helper columns are functional behavior.
- The public identity is `Yuan Jiang`; remove the personal email and unnecessary comment-author metadata.
- The public `Hanoi/` directory must contain only `hanoi.xlsx`, `README.md`, and `preview.png`.
- Do not push or publish to external communities in this implementation cycle.

---

## File Map

- Modify and rename: `Hanoi/Hanoi-1.0.xlsx` → `Hanoi/hanoi.xlsx` — executable workbook, internal Guide, formulas, and checks.
- Modify: `Hanoi/README.md` — public project landing page and download/inspection guidance.
- Replace: `Hanoi/Preview.png` → `Hanoi/preview.png` — clean non-zero-step release screenshot.
- Modify: `README.md` — stable Hanoi workbook and preview links plus verified compatibility wording.
- Reference: `docs/superpowers/specs/2026-07-29-hanoi-public-release-design.md` — approved requirements.
- Temporary only: `/private/tmp/hanoi-visual-baseline/` — before/after screenshots and range images; never commit.
- Temporary only: `/private/tmp/sanitize_hanoi_xlsx.py` — deterministic package sanitizer; never commit.

## Workbook Interfaces

- Raw controls: `Stage!B5` (`HanoiDisks`) and `Stage!B6` (`HanoiStep`).
- Effective controls: `HN_CLAMP_DISKS(value)` returns an integer from 3 through 6; `HN_EFFECTIVE_STEP(value,maxStep)` returns an integer from 0 through `maxStep`.
- Stable solver table: `Engine!A10:K74` with step, moving disk, source, destination, and six disk positions.
- Active step list: `_HanoiStepList` anchored at `Engine!M2`.
- Stage grid: `Stage!D5:AN13`; its dimensions and typography must remain unchanged.
- Formula checks: `Engine!D4:E9`, with aggregate status in `Engine!E9`.

---

### Task 1: Capture the Baseline and Record Workbook Contracts

**Files:**
- Read: `Hanoi/Hanoi-1.0.xlsx`
- Create temporarily: `/private/tmp/hanoi-visual-baseline/*`

**Interfaces:**
- Consumes: the connected `Hanoi-1.0.xlsx` Excel session.
- Produces: reproducible screenshots and recorded geometry for later comparison.

- [ ] **Step 1: Create the temporary evidence directory and record the current Git state**

Run:

```bash
mkdir -p /private/tmp/hanoi-visual-baseline
git status --short --branch
git rev-parse HEAD
```

Expected: branch `codex/hanoi-public-release`; only Excel's `Hanoi/~$Hanoi-1.0.xlsx` lock file may be untracked.

- [ ] **Step 2: Confirm the live workbook and close the ChatGPT pane**

Use Computer Use to confirm the Excel window title is `Hanoi-1.0.xlsx`, `Stage` is active, and the workbook path is inside this repository. Click only the ChatGPT pane's close button, then fetch fresh Excel state.

Expected: the full Stage layout is visible without the add-in pane and Excel remains at 100% zoom.

- [ ] **Step 3: Record the stage geometry and formulas**

Use Excel document control to read metadata and styles for:

```text
Stage!A1:AY20
Stage!D4:AN14
Engine!A1:M74
```

Record the row heights, column widths, hidden columns, font name/size, alignments, merged regions, data validations, and named formulas. The critical invariant is that `Stage!D5:AN13` retains identical geometry after edits.

- [ ] **Step 4: Capture three baseline states**

With the same Excel window size, 100% zoom, and active cell `A1`, capture the full Stage and `D4:AN14` for:

```text
disks=3, step=0
disks=4, step=7
disks=6, step=63
```

Save them under `/private/tmp/hanoi-visual-baseline/` using names such as `before-3-0-full.png` and `before-3-0-grid.png`. Restore `Stage!B5=3` and `Stage!B6=0` without saving unrelated UI state.

- [ ] **Step 5: Run the pre-edit workbook audit**

Search displayed values for:

```regex
#(REF!|VALUE!|NAME\?|CALC!|SPILL!|DIV/0!)
```

Inspect the XLSX package:

```bash
unzip -t Hanoi/Hanoi-1.0.xlsx
unzip -l Hanoi/Hanoi-1.0.xlsx
unzip -p Hanoi/Hanoi-1.0.xlsx xl/workbook.xml
```

Expected: no displayed formula errors and a valid archive; record the existing web-extension, comment, VML, and metadata parts as items to remove later.

---

### Task 2: Add Failing Robustness Checks, Then Harden the Named Formulas

**Files:**
- Modify: `Hanoi/Hanoi-1.0.xlsx`

**Interfaces:**
- Consumes: `Stage!B5`, `Stage!B6`, and the current named formulas.
- Produces: sanitized effective disk and step values used by all downstream formulas.

- [ ] **Step 1: Demonstrate the current disk-boundary mismatch**

Temporarily write `2` to `Stage!B5` through document control, read `Stage!B9`, and inspect the stage title.

Expected before the fix: the effective disk count is allowed below the documented minimum of 3. Restore `Stage!B5=3`.

- [ ] **Step 2: Demonstrate the current pasted-text failure**

Temporarily write `abc` to `Stage!B5` and read `Stage!B9`, `Stage!D4`, and `Engine!B5:B7`.

Expected before the fix: one or more dependent formulas report an error or cannot calculate a valid display. Restore `Stage!B5=3` and `Stage!B6=0`.

- [ ] **Step 3: Replace the disk sanitizer**

Update the workbook name `HN_CLAMP_DISKS` to:

```excel
=LAMBDA(n,LET(v,IFERROR(--n,3),MAX(3,MIN(6,ROUND(v,0)))))
```

This definition maps blanks, text, and formula errors to 3; rounds numeric values; and clamps the result to 3–6.

- [ ] **Step 4: Replace the effective-step sanitizer**

Update `HN_EFFECTIVE_STEP` to:

```excel
=LAMBDA(inputStep,maxStep,LET(v,IFERROR(--inputStep,0),MAX(0,MIN(maxStep,ROUND(v,0)))))
```

This definition maps blanks, text, and formula errors to 0; rounds numeric values; and clamps the result to the active maximum.

- [ ] **Step 5: Replace the progress helper**

Update `HN_PROGRESS_BAR` to:

```excel
=LAMBDA(step,maxs,bins,LET(fill,IF(maxs=0,0,ROUND(step/maxs*bins,0)),pct,IF(maxs=0,0,step/maxs),REPT("■",fill)&REPT("□",bins-fill)&"  "&TEXT(pct,"0%")))
```

Expected at three disks: step 0 shows `0%`; step 7 shows `100%`.

- [ ] **Step 6: Re-run the boundary matrix**

Test the following raw inputs and read the effective values from `Engine!B5:B7`:

| Raw disks | Raw step | Expected effective disks | Expected effective step |
|---:|---:|---:|---:|
| `2` | `0` | 3 | 0 |
| `7` | `0` | 6 | 0 |
| `4.6` | `8.6` | 5 | 9 |
| `abc` | `abc` | 3 | 0 |
| blank | blank | 3 | 0 |
| `4` | `99` | 4 | 15 |

Expected: no displayed formula errors. Restore `Stage!B5=3` and `Stage!B6=0`.

- [ ] **Step 7: Save and commit the formula hardening**

Save from Excel, wait for the lock file timestamp to settle, then run:

```bash
git add Hanoi/Hanoi-1.0.xlsx
git commit -m "fix: harden Hanoi formula inputs"
```

---

### Task 3: Build the Guide and Clarify the Stage

**Files:**
- Modify: `Hanoi/Hanoi-1.0.xlsx`

**Interfaces:**
- Consumes: the existing Stage visual language and named-formula groups.
- Produces: a discoverable internal user/developer guide and clearer Stage status text.

- [ ] **Step 1: Create and position the Guide sheet**

Create a visible worksheet named `Guide`, place it between `Stage` and `Engine`, and keep `Stage` active. Use the Stage dark green, medium green, pale green, and neutral text colors for a consistent workbook identity.

- [ ] **Step 2: Write the Guide content**

Merge each listed range and populate this exact English copy:

| Range | Content |
|---|---|
| `A1:H1` | `Hanoi Formula Step Demo` |
| `A2:H2` | `A formula-only Tower of Hanoi visualizer · Version 1.0.0` |
| `A4:H4` | `Quick Start` |
| `A5:H5` | `1. Open Stage. Choose 3–6 disks in B5.` |
| `A6:H6` | `2. Choose a step in B6 from 0 through Total Steps.` |
| `A7:H7` | `3. Follow Current Move, Disk Positions, the Motion Stage, and Keyframes.` |
| `A9:H9` | `Formula-Only Boundary` |
| `A10:H10` | `No VBA, macros, Office Scripts, custom functions, external runtimes, external data, or external workbook links.` |
| `A12:H12` | `Workbook Map` |
| `A13:H13` | `Stage — interactive controls, status, motion stage, legend, and keyframes.` |
| `A14:H14` | `Guide — usage, architecture, compatibility, and inspection notes.` |
| `A15:H15` | `Engine — solved states, active step list, named-formula integration, and release checks.` |
| `A17:H17` | `Formula Architecture` |
| `A18:H18` | `Stage inputs → HN_CLAMP_DISKS / HN_EFFECTIVE_STEP → Engine solved-state table` |
| `A19:H19` | `Engine states → query helpers → Stage title, status, keyframes, and cell-grid rendering` |
| `A21:H21` | `Named LAMBDA Groups` |
| `A22:H22` | `Input constraints — HN_CLAMP_DISKS, HN_EFFECTIVE_STEP, HN_MAX_STEP` |
| `A23:H23` | `Solver logic — HN_MOVE_DISK, HN_DISK_POS, HN_MOVE_ENDPOINT, HN_STEP_LIST` |
| `A24:H24` | `Lookup helpers — HN_MOVE_DISK_AT, HN_MOVE_FROM, HN_MOVE_TO, HN_POS_ROW` |
| `A25:H25` | `Display helpers — HN_STAGE_CELL, HN_STAGE_TITLE, HN_MOVE_LABEL, HN_PROGRESS_BAR, HN_POSITIONS` |
| `A27:H27` | `Compatibility` |
| `A28:H28` | `Verified on Microsoft Excel 365 Desktop. Excel 2024 and Excel for Web are unverified.` |
| `A30:H30` | `Inspect and Contribute` |
| `A31:H31` | `Inspect Engine and Excel Name Manager for the implementation. Report problems through GitHub Issues.` |
| `A32:H32` | `Created by Yuan Jiang with ChatGPT as an experiment in formula-only interactive software.` |

Use wrapped text, readable row heights, no clipped content, and no personal email address.

- [ ] **Step 3: Replace the misleading Stage status row**

Change `Stage!A13` from `Display Mode` to `Input Status`. Set `Stage!B13` to:

```excel
=LET(d,HN_CLAMP_DISKS($B$5),s,HN_EFFECTIVE_STEP($B$6,HN_MAX_STEP(d)),IF(AND($B$5=d,$B$6=s),"Inputs valid","Display adjusted to "&d&" disks / step "&s))
```

Keep the existing information strip below the motion stage, but update it to point readers to `Guide` for instructions and to explain that pasted invalid values are adjusted for display.

- [ ] **Step 4: Replace the Stage footer**

Remove the email address. Use the existing footer area for:

```text
Guide → see the Guide sheet for usage and formula architecture
v1.0.0 · Created by Yuan Jiang with ChatGPT · Formula-only
```

Keep the footer visually secondary to the stage.

- [ ] **Step 5: Verify the Stage did not shift**

Read the styles and geometry of `Stage!D4:AN14` and compare with Task 1. Confirm that no Stage grid row height, Stage grid column width, font, alignment, or hidden-helper setting changed.

- [ ] **Step 6: Save and commit the internal documentation**

Save from Excel, then run:

```bash
git add Hanoi/Hanoi-1.0.xlsx
git commit -m "docs: add Hanoi workbook guide"
```

---

### Task 4: Simplify Engine Documentation and Add Formula Checks

**Files:**
- Modify: `Hanoi/Hanoi-1.0.xlsx`

**Interfaces:**
- Consumes: `Engine!A10:K74`, `Engine!B5:B7`, and the named formulas from Task 2.
- Produces: `Engine!E5:E9` individual and aggregate PASS/FAIL results.

- [ ] **Step 1: Remove legacy notes and repeated comments**

Use Excel's native Clear Comments and Notes operation on the used ranges of `Stage` and `Engine`. Confirm that the visible cell contents and formulas remain unchanged.

Expected: no legacy comments remain, eliminating the repeated `Sheet2` text and comment-author metadata.

- [ ] **Step 2: Refresh the Engine introduction**

Unmerge and clear the old maintenance-note area `D4:L7`, leaving column M untouched. Replace the top explanatory copy in `A2:A3` with:

```text
Engine stores the solved-state table, active step list, and release checks used by Stage.
Stable contracts: A4:B7, A10:K74, and M1:M65. Inspect named HN_* formulas in Name Manager.
Change user-facing inputs on Stage. Do not insert rows or columns inside A10:K74.
```

Do not move or resize the solver table or active step list.

- [ ] **Step 3: Add the four invariant checks**

Create the block `Engine!D4:E9`:

```text
D4 Formula Checks        E4 Result
D5 Total steps           E5 =IF($B$6=2^$B$5-1,"PASS","FAIL")
D6 One disk per move     E6 =LET(n,$B$5,m,$B$6,p,TAKE($F$11:$K$74,m+1,n),c,--(DROP(p,1)<>TAKE(p,m)),IF(AND(BYROW(c,LAMBDA(r,SUM(r)))=1),"PASS","FAIL"))
D7 Distinct endpoints    E7 =LET(m,$B$6,src,TAKE($D$12:$D$74,m),dst,TAKE($E$12:$E$74,m),IF(AND((src<>"")*(dst<>"")*(src<>dst)),"PASS","FAIL"))
D8 Final state on C      E8 =LET(n,$B$5,m,$B$6,last,XLOOKUP(m,$A$11:$A$74,$F$11:$K$74),IF(AND(TAKE(last,,n)=3),"PASS","FAIL"))
D9 Overall               E9 =IF(COUNTIF($E$5:$E$8,"PASS")=4,"PASS","FAIL")
```

Format PASS in green and FAIL in red without adding scripts or macros.

- [ ] **Step 4: Test checks for every supported disk count**

For `Stage!B5` values 3, 4, 5, and 6, set `Stage!B6` to the maximum and read `Engine!E5:E9`.

Expected for every disk count: five `PASS` values and no Excel errors.

- [ ] **Step 5: Confirm solver contracts remain stable**

Read formulas at `Engine!A10:K12`, `Engine!A73:K74`, and `Engine!M1:M65`. Confirm no row or column insertion changed their addresses and `_HanoiStepList` still resolves to the spill anchored at `M2`.

- [ ] **Step 6: Save and commit the Engine cleanup**

Restore `Stage!B5=3`, `Stage!B6=0`, save, and run:

```bash
git add Hanoi/Hanoi-1.0.xlsx
git commit -m "test: add Hanoi formula invariants"
```

---

### Task 5: Run Functional and Visual Regression Tests

**Files:**
- Read: `Hanoi/Hanoi-1.0.xlsx`
- Create temporarily: `/private/tmp/hanoi-visual-baseline/after-*`

**Interfaces:**
- Consumes: the completed workbook content and Task 1 baselines.
- Produces: evidence that formulas and the cell-grid layout still behave correctly.

- [ ] **Step 1: Run the supported-state matrix**

For each disk count 3–6, test step 0, step 1, a middle step, and the maximum. At each state read the Stage title, current move, disk positions, keyframe current flag, Engine row for that step, and `Engine!E9`.

Expected: Stage and Engine agree; `Engine!E9` is `PASS`; no displayed formula errors exist.

- [ ] **Step 2: Repeat the three controlled screenshots**

With the same window dimensions, 100% zoom, and active cell as Task 1, capture:

```text
disks=3, step=0
disks=4, step=7
disks=6, step=63
```

Save `after-*-full.png` and `after-*-grid.png` files beside the baselines.

- [ ] **Step 3: Compare the cell-grid region**

Compare each before/after `D4:AN14` pair. The disk widths, peg centers, base line, row heights, column widths, font, and alignment must match. Differences are allowed only in deliberately changed copy outside the cell-grid region.

If a grid difference appears, stop and correct the workbook before continuing. A temporary duplicate Stage sheet may be created only for diagnosis and must be deleted afterward.

- [ ] **Step 4: Inspect the full-sheet layout**

Compare the full-sheet screenshots for panel balance, text wrapping, footer placement, legend alignment, keyframe alignment, hidden helpers, and clipping. Confirm the new Guide is readable at 100% zoom.

- [ ] **Step 5: Restore the release opening state**

Set `Stage!B5=3`, `Stage!B6=0`, select `Stage!A1`, keep Stage active, and save.

---

### Task 6: Sanitize the XLSX Package and Verify Reopening

**Files:**
- Modify: `Hanoi/Hanoi-1.0.xlsx`
- Create temporarily: `/private/tmp/sanitize_hanoi_xlsx.py`

**Interfaces:**
- Consumes: the saved workbook from Task 5.
- Produces: a valid XLSX package with no web extensions, legacy comments, VML note drawings, or unnecessary personal metadata.

- [ ] **Step 1: Close the workbook and preserve an exact temporary backup**

Close `Hanoi-1.0.xlsx` in Excel and confirm the lock file disappears. Copy the workbook to `/private/tmp/hanoi-visual-baseline/Hanoi-before-sanitize.xlsx`.

- [ ] **Step 2: Create a deterministic package sanitizer**

Create `/private/tmp/sanitize_hanoi_xlsx.py` with this ZIP/XML transform:

```python
from __future__ import annotations

import os
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

DROP_PREFIXES = ("xl/webextensions/",)
DROP_PARTS = {
    "xl/comments1.xml",
    "xl/comments2.xml",
    "xl/drawings/vmlDrawing1.vml",
    "xl/drawings/vmlDrawing2.vml",
}
DROP_REL_TYPES = {
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def should_drop(name: str) -> bool:
    return name in DROP_PARTS or name.startswith(DROP_PREFIXES)


def rewrite_xml(name: str, data: bytes) -> bytes:
    root = ET.fromstring(data)

    if name == "[Content_Types].xml":
        for child in list(root):
            part = child.attrib.get("PartName", "")
            extension = child.attrib.get("Extension", "")
            if (
                part.startswith("/xl/webextensions/")
                or part.startswith("/xl/comments")
                or extension.lower() == "vml"
            ):
                root.remove(child)

    elif name == "_rels/.rels" or (
        name.startswith("xl/worksheets/_rels/") and name.endswith(".rels")
    ):
        for child in list(root):
            if child.attrib.get("Type") in DROP_REL_TYPES:
                root.remove(child)

    elif name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
        for child in list(root):
            if local_name(child.tag) == "legacyDrawing":
                root.remove(child)

    elif name == "docProps/core.xml":
        for child in root:
            if local_name(child.tag) in {"creator", "lastModifiedBy"}:
                child.text = "Yuan Jiang"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sanitize(path: Path) -> None:
    fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
            temp_path, "w"
        ) as target:
            for info in source.infolist():
                if should_drop(info.filename):
                    continue
                data = source.read(info.filename)
                if (
                    info.filename in {"[Content_Types].xml", "_rels/.rels", "docProps/core.xml"}
                    or info.filename.startswith("xl/worksheets/")
                ) and info.filename.endswith((".xml", ".rels")):
                    data = rewrite_xml(info.filename, data)
                target.writestr(info, data)

        with zipfile.ZipFile(temp_path, "r") as check:
            bad_member = check.testzip()
            if bad_member is not None:
                raise RuntimeError(f"Corrupt ZIP member: {bad_member}")
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: sanitize_hanoi_xlsx.py WORKBOOK.xlsx")
    sanitize(Path(sys.argv[1]).resolve())
```

The transform removes matching `[Content_Types].xml` and relationship entries, removes `legacyDrawing` elements from worksheet XML, normalizes creator fields to `Yuan Jiang`, preserves other ZIP parts and metadata, validates the temporary output, and atomically replaces the source only on success.

- [ ] **Step 3: Run the sanitizer and scan the package**

Run:

```bash
python3 /private/tmp/sanitize_hanoi_xlsx.py Hanoi/Hanoi-1.0.xlsx
unzip -t Hanoi/Hanoi-1.0.xlsx
unzip -l Hanoi/Hanoi-1.0.xlsx
```

Expected: no `xl/webextensions`, `comments*.xml`, `vmlDrawing*.vml`, or invalid ZIP member.

- [ ] **Step 4: Verify prohibited artifacts are absent**

Run package searches for:

```text
vbaProject
officeScripts
externalLinks
connections.xml
webextensions
LABS_GENERATIVEAI
wa200010215
wa200003696
Sheet2
sleepdrop@gmail.com
```

Expected: zero matches in package paths and extracted XML/text.

- [ ] **Step 5: Reopen in Excel and check for repair warnings**

Open the sanitized workbook. Expected: Excel opens it normally without a repair dialog. Confirm `Stage`, `Guide`, and `Engine` exist in the correct order, Stage is active at 3 disks and step 0, Formula Checks report PASS, and no displayed formula error is present.

- [ ] **Step 6: Save, close, and re-scan**

Save once in Excel, close the workbook, and repeat the ZIP and prohibited-artifact scans. If Excel reintroduces a web extension because an add-in pane was reopened, close the pane, sanitize again, and repeat this step.

- [ ] **Step 7: Commit the sanitized workbook**

Run:

```bash
git add Hanoi/Hanoi-1.0.xlsx
git commit -m "chore: sanitize Hanoi workbook package"
```

---

### Task 7: Publish the Stable Local Artifact and Documentation

**Files:**
- Rename: `Hanoi/Hanoi-1.0.xlsx` → `Hanoi/hanoi.xlsx`
- Modify: `Hanoi/README.md`
- Rename/replace: `Hanoi/Preview.png` → `Hanoi/preview.png`
- Modify: `README.md`

**Interfaces:**
- Consumes: the verified sanitized workbook.
- Produces: the final public project directory and repository navigation.

- [ ] **Step 1: Rename the workbook**

Run:

```bash
git mv Hanoi/Hanoi-1.0.xlsx Hanoi/hanoi.xlsx
```

Expected: no workbook with a version number remains in `Hanoi/`.

- [ ] **Step 2: Rewrite the Hanoi README**

Include these exact sections in concise English:

```text
Title and one-sentence formula-only description
Preview
Try It / Download
Quick Start
What It Demonstrates
Workbook Map
Formula Architecture
Requirements and Compatibility
Inspecting the Implementation
Feedback and License
```

Link the download to `hanoi.xlsx`, the image to `preview.png`, and feedback to the repository's GitHub Issues page. State that Excel 365 Desktop is verified while Excel 2024 and Excel for Web are unverified.

- [ ] **Step 3: Update the repository README**

Change the Hanoi download link to `Hanoi/hanoi.xlsx`. Keep the project description concise and consistent with the Hanoi README. Do not claim compatibility that was not tested.

- [ ] **Step 4: Capture the public preview**

Open `hanoi.xlsx`, close the ChatGPT pane, set 4 disks and step 7, select `Stage!A1`, and keep 100% zoom. Capture a clean Excel-window screenshot with the entire Stage visible and no personal information. Save it as `Hanoi/preview.png` and remove `Hanoi/Preview.png`.

- [ ] **Step 5: Restore and save the opening state**

Set 3 disks and step 0, select `Stage!A1`, save, close, and confirm no lock file remains.

- [ ] **Step 6: Verify the public directory**

Run:

```bash
find Hanoi -maxdepth 1 -type f -print | sort
```

Expected exactly:

```text
Hanoi/README.md
Hanoi/hanoi.xlsx
Hanoi/preview.png
```

- [ ] **Step 7: Commit the release files**

Run:

```bash
git add README.md Hanoi/README.md Hanoi/hanoi.xlsx Hanoi/preview.png
git commit -m "release: prepare Hanoi workbook for public use"
```

---

### Task 8: Final Release Verification

**Files:**
- Verify: `Hanoi/hanoi.xlsx`
- Verify: `Hanoi/README.md`
- Verify: `Hanoi/preview.png`
- Verify: `README.md`

**Interfaces:**
- Consumes: all completed Hanoi release artifacts.
- Produces: evidence-backed approval that the branch is ready for review, but not pushed.

- [ ] **Step 1: Reopen the final filename and run smoke tests**

Open `Hanoi/hanoi.xlsx` in Excel. Test 3 disks at step 0, 4 disks at step 7, and 6 disks at step 63. Confirm Stage/Engine agreement, Formula Checks PASS, keyframe highlighting, and valid layout.

- [ ] **Step 2: Run the final formula-error scan**

Search displayed values across the workbook using:

```regex
#(REF!|VALUE!|NAME\?|CALC!|SPILL!|DIV/0!)
```

Expected: zero matches.

- [ ] **Step 3: Run the final package audit**

Confirm archive integrity and absence of VBA, scripts, connections, external links, web extensions, legacy comments, personal email, stale `Sheet2`, and temporary sheets.

- [ ] **Step 4: Verify documentation and links**

Check that every local link in `README.md` and `Hanoi/README.md` resolves with exact filename case. Confirm the preview matches the current workbook and all copy is English.

- [ ] **Step 5: Check Git cleanliness and review the diff**

Run:

```bash
git status --short --branch
git diff --check HEAD~1
git log --oneline --decorate -8
```

Expected: no lock files or unintended artifacts, no whitespace errors, and only the planned Hanoi release changes plus the approved design/plan documents.

- [ ] **Step 6: Stop before external publication**

Report the verification evidence and branch state to the user. Do not push, open a pull request, create a GitHub release, or post to Hacker News/Reddit without a later explicit instruction.
