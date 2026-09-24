# Round 03 — Fixed-clue integrity

Date: 2026-09-24. Base: local `main` at `fca9d48`; production workbook SHA-256 `ca7457b4ce35792b63fa33b324f77d2871c7eddb0d66fafe303d0c5d66beb637`. Remote freshness remains unverified because the remote check could not connect in the earlier round. Prior review commits were already on local `main`, so this round began without a preliminary commit.

## Boundary and reproduction

F1 only: prevent accidental fixed-clue deletion across the five player sheets while keeping player entries and Coach mode editable. Round 02 had reproduced the defect in a clean Excel copy: Delete cleared `05 Master!S17`, a gray clue formula with value 8, and produced `-1 answers entered`. The read-only regression check failed on the source workbook: all five pages were unprotected, 261 blank player centers were locked, and all five Coach selectors were locked.

## Smallest verified fix

In the clean temporary copy `/private/tmp/sudoku-review-round2-clue-repro/sudoku-clue-repro.xlsx`, Microsoft Excel unlocked only the blank player centers and each `AS11:AZ11` merged Coach selector, then enabled native protection on the five player sheets without a password. The 261 input centers remain the only editable board cells. The five hidden engines and `_Tests` remain unprotected and unchanged. The protected sheets show lock icons on their tabs.

An attempted native Excel clear of `05 Master!S17` was rejected with the protected-sheet error; its original formula and value 8 remained. After save and reopen, pressing **Delete** and **Backspace** on that clue each showed Excel's protected-sheet alert, and the formula and value still remained. In the reopened copy, typing `1` into `05 Master!K26` changed status to one answer entered and remaining to 57; Delete cleared it and restored zero answers and 58 remaining.

Through the connected Excel session, one correct answer was also entered and cleared on each page: Easy `S17=5`, Medium `G5=3`, Hard `K23=6`, Expert `AE14=2`, Master `K26=1`. Every page reached one answer entered, its expected remaining count, and zero conflicts, then returned to its clean count with Coach Off. Master `AS11=Hint` remained editable and produced `Hidden single in column · R8C3 = 1`; returning to Off restored the clean state. This verifies protection did not block the existing candidate/Coach calculation path for the tested edit.

## Save, regression, and production state

The temporary copy was saved, closed, and reopened in Excel. The strengthened [audit_clue_protection.py](audit_clue_protection.py) passed on that saved file and still failed on the old production file before replacement. It now also rejects unexpected unlocked cells. A semantic comparison with the source found 13 sheets and 78 defined names unchanged, no cell-content or formula changes, no visual style changes, and no merged-range, validation-count, or conditional-formatting-count changes. The 301 protection-style changes are exactly 261 player centers plus eight merged Coach cells on each of five pages. ZIP integrity passed.

The verified saved copy was then copied to `Sudoku/sudoku.xlsx`. Production SHA-256 is now `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`, byte-identical to the saved test copy. Production readback shows all five puzzles clean (remaining 51/45/53/54/58, zero conflicts, Coach Off), `_Tests` **131 PASS / 0 FAIL / 18 SKIP**, `E145=PASS`, workbook version `1.0.0`, and the candidate display audit passes all 405 squares. The protection audit and ZIP integrity also pass. No candidate-chain migration or formula change was made. Existing untracked experiments were untouched.

## Overall progress and next bounded round

R1 and F1 are complete. R2 has candidate-display checks on all five pages and clean dead-end checks on four; Master dead-end and clear recovery remain open. R3–R6 remain open. Next, finish only the Master part of R2 using a fresh copy of this protected production workbook: reproduce a no-duplicate dead end, clear each test answer, confirm candidates/status/Coach return to baseline, and save/reopen. Then move to Coach modes in a later round.
