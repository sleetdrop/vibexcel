# Round 06 — Cross-sheet isolation (complete)

Date: 2026-09-24. Base: `main` at `51a704a`; production `Sudoku/sudoku.xlsx` SHA-256 `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`. Tested a fresh byte-identical copy in Microsoft Excel 365 for macOS: `/private/tmp/sudoku-review-round6-isolation-20260924/sudoku-isolation.xlsx`. Saved a clean Excel-recalculated `baseline.xlsx` before editing, then saved snapshots after each source edit. Comparisons below cover every cached cell of all five player pages and their five matching engines, including status, candidates, Coach output, and chain caches. Formula/value storage on the unaffected sheets was also compared.

## Easy as the only source

Entered correct `01 Easy!S17=5` and set `AS11=Hint`. Native Excel showed `IN PROGRESS · 1 answers entered`, 50 remaining, zero conflicts, and `Naked single · R5C2 = 2`; the changed candidate display and source marker were visible. The other four pages stayed at zero answers, Off, blank guidance, zero conflicts, and their original remaining counts 45/53/54/58. The saved snapshot differed from baseline in **14 Easy page cells and 22 `_Engine` cells**; each of the other four player pages and engines had **zero cached-value differences**. No formula/value storage changed on those eight unaffected sheets.

## Master as the only source

Cleared Easy `S17`, returned Easy to Off, entered correct `05 Master!K26=1`, and selected Trace. Native Excel showed `IN PROGRESS · 1 answers entered`, 57 remaining, zero conflicts, and `Technique — Coach reveal | Pattern — R1C2 = 6` followed by the explicit certified-solution disclosure. Easy, Medium, Hard, and Expert showed zero answers, Off, blank guidance, zero conflicts, and remaining counts 51/45/53/54. The saved snapshot differed from baseline in **13 Master page cells and 17 `_Engine Master` cells**; the other four player pages and engines again had **zero cached-value differences**, with no formula/value storage changes on those eight sheets.

## Restore and regression

Cleared Master `K26`, set Off, saved, closed, and reopened the copy in Excel. The reopened Master page showed zero answers, 58 remaining, zero conflicts, and Off. All cached cells of the five player pages and five engines matched the clean Excel baseline exactly; their formula/value storage also matched. The final saved `_Tests` result was **131 PASS, 0 FAIL, 18 SKIP**. The candidate audit passed all **405** squares and the clue-protection audit passed all five pages. Production SHA-256 remained unchanged.

During either intentional player edit, `_Tests` showed 128 PASS, 3 FAIL, 18 SKIP. The three failures were initial-state assertions at rows 132, 133, and 135: all modes start Off, all boards start without player entries, and remaining counters match initial clues. They returned to PASS after reset; they do not indicate cross-sheet leakage. R6 should explicitly reconcile these passive test assumptions with live gameplay.

No isolation defect was reproduced, so the production workbook and regression formulas were not changed. R1–R4 and F1 are complete; R5 and R6 remain open. Next bounded round: R5 save/reopen behavior for an **in-progress** game, including answer, mode, candidates, status, manual clearing, and fresh-copy reset. Existing untracked candidate-chain experiments were left untouched.
