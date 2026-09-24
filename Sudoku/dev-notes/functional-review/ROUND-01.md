# Round 01 — Entry, duplicate conflict, and clear recovery

Date: 2026-09-24. Base: local `main` and `origin/main` at `d88e72c2435984418ea6073d08e84fc5b26b7e1e`. The GitHub remote could not be reached in this environment, so remote freshness was not independently verified. Production workbook SHA-256: `ca7457b4ce35792b63fa33b324f77d2871c7eddb0d66fafe303d0c5d66beb637`. Native testing used `/private/tmp/sudoku-review-round1-20260924/sudoku-review.xlsx`, copied from that exact file.

## Scope and observations

All entries below were made in the visible green-underlined player cells in Microsoft Excel 365 for macOS. Each wrong digit duplicated an existing clue. The entered square stopped showing candidates; clearing restored its candidate display, empty input, `IN PROGRESS · 0 answers entered`, zero conflicts, and the original remaining count.

| Difficulty | Wrong entry | Conflict count | Remaining while entered → after clear |
| --- | --- | ---: | ---: |
| Easy | `K5 = 5` | 2 | 50 → 51 |
| Medium | `C5 = 6` | 3 | 44 → 45 |
| Hard | `C5 = 9` | 3 | 52 → 53 |
| Expert | `C5 = 3` | 2 | 53 → 54 |
| Master | `G5 = 1` | 3 | 57 → 58 |

Easy `S17 = 5` was also entered as a correct answer: status stayed `IN PROGRESS`, conflicts stayed 0, and remaining fell from 51 to 50. Delete restored the starting state. Both Delete and Backspace cleared player input in native Excel; one initial Hard Backspace attempt left the value in place, but an immediate Delete and a fresh Backspace reproduction both cleared it. The isolated attempt is not a confirmed formula defect.

After all entries were cleared, the temporary copy was saved, closed, and reopened in Excel. Cached readback showed all five pages at Coach `Off`, blank tested input cells, 0 conflicts, and remaining counts 51/45/53/54/58. Hidden `_Tests` reported **131 PASS / 0 FAIL / 18 SKIP**, including `E145 = PASS`; ZIP integrity passed. The 18 skipped checks were not executed in this round.

## Findings and change decision

No stable functional defect was found within this round's boundary; no workbook formulas or regression cells were changed. The repository README and architecture document still reported the historical 79-pass release count, whereas the workbook `README!D45` and reopened `_Tests` show 131 PASS and 18 SKIP. The documents were updated to identify the current workbook baseline and its already-connected `SDKP_*` candidate/Coach path. Workbook version `README!D54 = 1.0.0` and build date `README!D55 = 2026-07-27` remain consistent with the repository README; neither was inferred from the Git commit date.

## Overall progress and next round

The five-engine candidate-chain migration is already complete; this review did not redo it. **1 of 6** bounded player-path ranges in [CHECKLIST.md](CHECKLIST.md) is complete. Next: R2, independently check candidate legality and non-duplicate dead ends after legal-looking wrong entries, including clear recovery. Coach modes, cross-sheet isolation, in-progress save/reopen, and on-demand hidden checks remain open.

Production `Sudoku/sudoku.xlsx` was not opened for editing and must retain the SHA-256 above. Existing untracked candidate-chain experiments are outside this round and must be preserved.
