# Round 07 — Save, reopen, and reset (complete)

Date: 2026-09-24. Base: `main` at `5ec0335`. Production `Sudoku/sudoku.xlsx` SHA-256 `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`. Tested a fresh byte-identical copy in Microsoft Excel 365 for macOS: `/private/tmp/sudoku-review-round7-persistence-20260924/sudoku-persistence.xlsx`. A clean Excel-saved `baseline.xlsx` was preserved before player input.

## In-progress save and reopen

Using the native player UI, entered `05 Master!K26=1` and selected `AS11=Trace`. The player page showed `IN PROGRESS · 1 answers entered`, 57 remaining, zero conflicts, and `Technique — Coach reveal | Pattern — R1C2 = 6` with an explicit certified-solution source disclosure. The saved candidate display changed around R8C3: `J25/K25/N25/R25/V25/L26/K27` changed or cleared, and the matching engine candidate projection and chain cache changed. The saved `_Tests` had **128 PASS, 3 FAIL, 18 SKIP**; the three FAILs were the expected initial-state assertions at rows 132, 133, and 135 while a player entry and Trace mode were active.

Saved, closed, and reopened the in-progress copy in native Excel. The screen still showed the answer, 57 remaining, Trace mode, Coach disclosure, and changed candidates. Saved again after reopening. A cached-value comparison between the before/after reopen saves found **zero differences across all 13 sheets**, including `_Tests`, all five player pages, and all five engines.

## Manual clear and clean state

Selected editable `K26` in the grid and pressed Delete. Excel changed to zero answers and 58 remaining. Entered `Off` in the Coach selector through the player UI; the guidance cleared. Saved and closed the clean copy, then reopened it. Native Excel showed zero answers, 58 remaining, zero conflicts, Off, and the initial candidate display. The clean saved file had **131 PASS, 0 FAIL, 18 SKIP**. Its cached values matched the Excel-saved baseline at every cell of all 13 sheets. Independent audits passed all **405** candidate squares and protection on all five player pages. A separate fresh reset copy made directly from production was byte-identical to production (same SHA-256).

## Final post-reopen read

After the Mac was unlocked, the already reopened clean copy was saved again in Excel. The saved `_Tests` still had **131 PASS, 0 FAIL, 18 SKIP**. Every cached cell on all 13 sheets matched the clean pre-reopen save (zero differences). Master `K26` was blank, `AS11` was Off, status showed zero answers, remaining was 58, conflicts were zero, and Coach guidance was blank. No persistence defect was reproduced, so no production workbook or regression formula change was needed. Production SHA-256 remained unchanged. Existing untracked candidate-chain experiments were not touched.

R1–R5 and F1 are complete; R6 remains open. Next bounded round: hidden regression and release contracts, including on-demand uniqueness and solver checks where resource use allows, formula-error and validation inspection, and reconciliation of workbook README counts against the saved workbook.
