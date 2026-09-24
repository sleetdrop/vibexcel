# Round 04 — Master dead-end and clear recovery

Date: 2026-09-24. Base: local `main` at `3842391`; production workbook SHA-256 `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`. F1 had been committed first. The previous R2 Master example was invalid because a clue had been accidentally cleared; this round used a fresh byte-identical protected copy at `/private/tmp/sudoku-review-round4-master-20260924/sudoku-master.xlsx`.

## Exact player path

An independent row, column, and box calculation on the clean Master puzzle found two individually legal entries that together leave `R2C3` (`K8`) with no legal digit: `G5=6`, then `K5=4`. In Excel, the first entry showed one answer entered, 57 remaining, zero conflicts. The second showed **DEAD END · 2 answers entered**, 56 remaining, **zero conflicts**. `K8` was blank and all eight visible candidate positions around it were blank. A saved readback confirmed both inputs, the dead-end status, zero conflicts, and the empty dead cell. [audit_candidates.py](audit_candidates.py) passed all 405 squares in the dead-end copy, including Master chain projection and basic legality checks.

Clearing `K5` through Excel's Delete key immediately restored **IN PROGRESS · 1 answers entered** and 57 remaining. Clearing `G5` restored **IN PROGRESS · 0 answers entered** and 58 remaining. The copy was saved, closed, and reopened. The five visible game pages then had **zero cached cell-value differences** from the clean production workbook; each showed Coach Off and zero conflicts. The candidate audit again passed all 405 squares, the protection audit passed, and hidden `_Tests` returned **131 PASS / 0 FAIL / 18 SKIP**, including `E145=PASS`.

## Result and overall progress

No defect was found in this bounded Master path, and the production workbook was not changed. R2 is now complete alongside R1 and F1. R3–R6 remain open. Next round: Coach modes across all five difficulties, starting from a fresh protected copy, with Off → Hint → Trace → Verify → Off and explicit checks after correct and erroneous entries. Verify may invoke expensive global search, so assess it one difficulty at a time and record any resource limit that affects use.
