# Round 02 — Candidate display and non-duplicate dead ends

Date: 2026-09-24. Base: local `main` at `ec126d7`; production workbook SHA-256 `ca7457b4ce35792b63fa33b324f77d2871c7eddb0d66fafe303d0c5d66beb637`. Remote freshness remains unverified because `git ls-remote origin` could not connect. Testing used a fresh `/private/tmp/sudoku-review-round2-20260924/sudoku-review.xlsx` copied from production and opened in Microsoft Excel 365 for macOS.

## Candidate contract and results

The active `SDKP_Chain` can eliminate digits beyond direct row/column/box exclusions. The independent check therefore requires each shown candidate to equal its engine projection and be a subset of basic legal digits. It also checks all eight visible candidate positions and the center value of each logical square. [audit_candidates.py](audit_candidates.py) passed all **405 squares** on the clean copied file and after one correct player input on each difficulty:

| Difficulty | Correct input | Saved state |
| --- | --- | --- |
| Easy | `S17 = 5` | In Progress, 0 conflicts |
| Medium | `G5 = 3` | In Progress, 0 conflicts |
| Hard | `K23 = 6` | In Progress, 0 conflicts |
| Expert | `AE14 = 2` | In Progress, 0 conflicts |
| Master | `K26 = 1` | In Progress, 0 conflicts |

Each page showed one answer entered. The five inputs were cleared before dead-end testing. The audit also passed all 405 squares on the saved dead-end copy, which establishes candidate display consistency for those saved states; it does not establish that every chain elimination is logically sound.

## No-duplicate dead ends

An independent row/column/box calculation selected entries that were locally legal when typed but left another empty square with no basic legal candidate. Native Excel showed `DEAD END` and 0 direct conflicts in these four intact puzzles; saved readback confirmed the indicated empty dead cells.

| Difficulty | Entry sequence | Empty dead cell | Dashboard answers |
| --- | --- | --- | ---: |
| Easy | `AI5 = 4` | R7C9 | 1 |
| Medium | `C5 = 3` | R1C2, R5C1 | 1 |
| Hard | `K5 = 6` | R7C3 | 1 |
| Expert | `C5 = 6`, then `S5 = 7` | R1C4 | 2 |

The Master sequence `K5 = 4`, `AE8 = 6` was not accepted as a clean-puzzle result. Its saved copy showed both inputs and `DEAD END`, but the answer count was only 1 because fixed clue `S17 = 8` had been cleared during an earlier multi-sheet input-clear operation. The affected copy is retained under `/private/tmp`; Master dead-end and clear recovery remain open.

## Confirmed fixed-clue defect

A second fresh copy, `/private/tmp/sudoku-review-round2-clue-repro/sudoku-clue-repro.xlsx`, started byte-identical to production. In native Excel, selecting gray fixed clue `05 Master!S17` showed the original engine-linked formula and value 8. Pressing Delete cleared the formula without a validation warning; remaining rose from 58 to 59, and status displayed `IN PROGRESS · -1 answers entered`. Undo restored the formula. Static inspection found all five player sheets unprotected, with all green-underlined input centers and Coach selectors still marked locked. The fixed clues use custom `FALSE` data validation, which does not intercept Delete.

[audit_clue_protection.py](audit_clue_protection.py) is a new read-only regression gate. It currently exits 1 on production: five sheets are unprotected, 261 player centers are locked, and five Coach selectors are locked. A viable fix must protect the sheets, unlock only player centers and Coach selectors, and pass native input, mode, save/reopen, and visual checks. Selecting all 59 Master inputs as one non-contiguous UI range made Excel stop responding, so the protection change was **not** applied to production or claimed verified. The failed GUI path should not be repeated unchanged.

## Progress and next bounded round

R1 remains complete. R2 has passed baseline and changed-board candidate display checks on all five pages, plus four clean dead-end cases; it is **not complete** because Master and clear recovery remain open. F1 fixed-clue integrity is the next high-risk boundary. No formula or binary workbook change was made in this round. The repository README and architecture note now warn about Delete clearing fixed clues. Existing untracked candidate-chain experiments remain untouched.
