# Round 05 — Coach modes (in progress)

Date: 2026-09-24. Base: `main` at `6fd70e1`, pushed to `origin/main` after a successful fetch and fast-forward check (`d88e72c..6fd70e1`). Production workbook SHA-256 `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`. Testing began in a fresh byte-identical copy, `/private/tmp/sudoku-review-round5-coach-20260924/sudoku-coach.xlsx`, connected to Microsoft Excel 365 for macOS. Production was not opened for editing.

## Clean-board mode cycle

Each page was switched Off → Hint → Trace → Verify → Off. In all five pages, Hint matched the matching engine's `B88`; Trace named the technique and supplied a separate action line; Verify displayed `VIABLE · At least one completion remains` and asked the player to return to Off; Off cleared the guidance. The clean Hint targets were:

| Page | Clean Hint target | Technique |
| --- | --- | --- |
| Easy | R5C5 = 5 | Naked single |
| Medium | R1C2 = 3 | Naked single |
| Hard | R7C3 = 6 | Naked single |
| Expert | R4C8 = 2 | Hidden single in row |
| Master | R8C3 = 1 | Hidden single in column |

An independent read-only check of the saved production board found every target blank, its digit in the engine candidate set, legal by row/column/box, and equal to the certified solution. Master Verify took about 10.8 seconds for the mode change and readback in the connected session; it completed without an error.

## Changed-board checks completed

- Master `K26=1` advanced Hint to `Coach reveal · R1C2 = 6`, explicitly saying the value comes from the certified solution because no supported deduction was available. Trace repeated that source disclosure. Verify remained `VIABLE`. Adding duplicate `G5=1` produced `CONFLICT · 2 answers entered`, three conflicts, and `Resolve the conflict before using Coach` in Hint, Trace, and Verify. A later read confirmed both entries cleared, Coach Off, zero conflicts, and 58 remaining after a reset command timed out at the transport layer.
- Easy `S17=5` advanced Hint to Naked single `R5C2=2`; duplicate `K5=5` produced two conflicts and the explicit conflict guidance. Reset returned to clean status and Off.
- Medium `G5=3` advanced Hint to Naked single `R1C3=5`; duplicate `C5=6` produced three conflicts and explicit conflict guidance. Reset command succeeded.

## Interruption and remaining work

The Mac locked during the round. The connected Excel session continued through the mode cycles and the changed-board checks above, but the Hard correct/error command timed out, its follow-up state read also timed out, and the Excel session then disappeared. The Hard state after that command is unknown. Expert changed-board behavior, UI marker colors/roles, and this round's save/reopen and regression readback are not yet verified. Do not interpret the transport timeout as a formula defect or mark R3 complete. The current temporary copy may contain unsaved changes; use a **new** production copy after the Mac is unlocked to finish the remaining checks.

Overall, R1, R2, and F1 remain complete; R3 is in progress; R4–R6 remain open. Existing untracked candidate-chain experiments were not touched.
