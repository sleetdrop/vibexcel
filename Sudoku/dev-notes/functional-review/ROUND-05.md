# Round 05 — Coach modes (complete)

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

## Resume after the Mac was unlocked

The interrupted copy was closed without saving. A new byte-identical production copy, `/private/tmp/sudoku-review-round5-coach-resume-20260924/sudoku-coach-resume.xlsx`, was opened in native Excel. This avoids drawing conclusions from the timed-out command's unknown state.

- Hard: correct `K23=6` gave `IN PROGRESS · 1 answers entered`, 52 remaining, zero conflicts, and a Naked single Hint `R8C1 = 8`. Duplicate `C5=9` gave `CONFLICT · 2 answers entered`, 51 remaining, three conflicts, and `Resolve the conflict before using Coach.` Clearing both and selecting Off restored 53 remaining, zero conflicts, and blank guidance.
- Expert: correct `AE14=2` gave `IN PROGRESS · 1 answers entered`, 53 remaining, zero conflicts, and a Hidden single in row Hint `R4C7 = 9`. Duplicate `C5=3` gave `CONFLICT · 2 answers entered`, 52 remaining, two conflicts, and the same explicit block. Clearing both and selecting Off restored 54 remaining, zero conflicts, and blank guidance.
- On clean Hint in each difficulty, a scan of the 81 grid corner cells found exactly one `●` source glyph and no `▼` target or `✖` conflict glyph: Easy R5C5, Medium R1C2, Hard R7C3, Expert R4C8, Master R8C3. Native Excel showed the Master glyph and Coach text together. The hidden coordinate-role tests cover source and target parsing, including the `▼` target for a pointing elimination; the current clean player boards did not display an elimination target.
- On Master, entered six correct values used by the reachable advanced fixture: `K26=1`, `G5=6`, `K8=4`, `K17=3`, `K23=6`, `C8=5`. The current chain-connected player Hint was instead `Hidden single in column · R6C7 = 1`, with one `●` at that source and no target glyph. The older `_Tests` player-trace fixture uses a separate reasoning function and therefore is not evidence that the published chain must show pointing at that exact state. Replacing editable `G5` with duplicate `5` produced two visible `✖` conflict glyphs, zero source/target glyphs, and the explicit conflict message; native Excel showed the same state. An attempted edit of the fixed clue `C5` was rejected by worksheet protection, as intended.

All six Master entries were cleared and all five modes set to Off. A live read showed zero answers, zero conflicts, blank Coach guidance, and remaining counts 51/45/53/54/58. The fresh copy was saved, closed, and reopened in Excel. Its saved visible cell values match production on all five player sheets (zero differences). The saved hidden suite has **131 PASS, 0 FAIL, 18 SKIP**. The independent candidate audit passed all **405** squares, and the protection audit passed all five pages. Production SHA-256 remained `1de172cba70d6d784eaa3ac77280d40d272bd0f86a9ffa4f863d9eb14d80e852`.

No formula defect was reproduced, so no production workbook change or new regression formula was warranted. R1, R2, F1, and R3 are complete; R4–R6 remain open. Next bounded round: R4 cross-sheet isolation, comparing visible pages and their matching engine caches before and after edits on two source difficulties. Existing untracked candidate-chain experiments were not touched.
