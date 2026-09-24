# Whole-workbook functional review checklist

The executable `Sudoku/sudoku.xlsx` is the source of truth. Run each round on a fresh temporary copy in Microsoft Excel 365 Desktop. Mark an item complete only after observing the player-facing result and, where relevant, the saved workbook. Preserve the production workbook until a defect has been reproduced and a narrow fix has passed regression checks.

## Player paths

- [x] **R1 — Entry, conflict, clear, and recovery across all five difficulties.** For each page, enter a duplicate in an editable center. Check conflict highlighting, game status, conflict and remaining counts, candidate display in the entered square, and restoration after clearing. Check one correct entry and its reversal. See [ROUND-01.md](ROUND-01.md).
- [x] **R2 — Candidate legality and local dead ends.** All five pages passed baseline and legal-edit candidate checks, then non-duplicate dead-end and clear recovery checks. Chain eliminations may legitimately narrow the independent row, column, and box candidate set. See [ROUND-02.md](ROUND-02.md) and [ROUND-04.md](ROUND-04.md).
- [x] **F1 — Fixed-clue integrity.** Native protection now prevents Delete/Backspace from removing gray clue formulas while green-underlined centers and the Coach selector remain editable. All five pages, save/reopen, [audit_clue_protection.py](audit_clue_protection.py), and the existing workbook suite passed. See [ROUND-03.md](ROUND-03.md).
- [x] **R3 — Coach modes.** All five pages passed Off → Hint → Trace → Verify → Off, target/digit legality, changed-board and conflict recovery, fallback disclosure, and source/conflict marker checks in native Excel. The hidden suite covers elimination target roles; the published boards did not surface a target during this round. See [ROUND-05.md](ROUND-05.md).
- [x] **R4 — Cross-sheet isolation.** Easy and Master each served as the only edited source. In both saved states, every cached cell on the other four player pages and four engines matched the clean Excel baseline; native status and Coach output agreed. Clearing both edits restored all ten sheets. See [ROUND-06.md](ROUND-06.md).
- [x] **R5 — Save, reopen, and reset.** In-progress Master answer, Trace, candidates, status, and hidden state persisted across save/reopen; all 13 cached sheets matched before/after. Native Delete and Off restored the clean state, which also persisted after reopen; a fresh production copy matched byte-for-byte. See [ROUND-07.md](ROUND-07.md).
- [ ] **R6 — Hidden regression and release contracts.** Run passive `_Tests` and selected on-demand checks, including uniqueness and solver status where resource use allows. Inspect formula errors, broken references, validation, and visual gates. Reconcile counts with the workbook `README`.

## Per-round record

Record the exact input and expected result, observed Excel behavior, any defect and root cause, the smallest fix and meaningful regression check if needed, save/reopen evidence, production workbook hash and Git status, overall progress, and the next bounded range. Do not count an unrun checklist item as verified.
