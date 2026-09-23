# Master candidate-chain production-copy pilot

## Purpose and boundary

This is the exact change and acceptance list for a one-engine pilot in a fresh
copy of `Sudoku/sudoku.xlsx`. It is not authorization to edit the production
workbook. Start from a byte-identical copy of production SHA-256
`f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.

The verified source prototype is
`sudoku-candidate-chain-structured-slice.xlsx`, SHA-256
`72d149befad45466a869d3c93b6dd4c9f2cb177a0e14d24d8ce143b704f45d5c`.
Its native test baseline is **131 PASS, 0 FAIL, 18 SKIP**.

## Exact pilot changes

### 1. Add the shared workbook names

Copy these 27 workbook-scoped names and their formulas from the verified
prototype. Do not alter any existing production name; package comparison
confirms that the prototype changes no pre-existing defined-name formula.

1. `SDKP_AdvancedHint`
2. `SDKP_Apply`
3. `SDKP_ApplyStep`
4. `SDKP_BasicHint`
5. `SDKP_Chain`
6. `SDKP_ChainLoop`
7. `SDKP_ClaimingCol`
8. `SDKP_ClaimingColStep`
9. `SDKP_ClaimingRow`
10. `SDKP_ClaimingRowStep`
11. `SDKP_Contradiction`
12. `SDKP_ForcedConflict`
13. `SDKP_HasProgress`
14. `SDKP_HiddenBox`
15. `SDKP_HiddenCol`
16. `SDKP_HiddenRow`
17. `SDKP_Naked`
18. `SDKP_NakedPairBox`
19. `SDKP_NakedPairBoxStep`
20. `SDKP_NakedPairCol`
21. `SDKP_NakedPairColStep`
22. `SDKP_NakedPairRow`
23. `SDKP_NakedPairRowStep`
24. `SDKP_PointingCol`
25. `SDKP_PointingColStep`
26. `SDKP_PointingRow`
27. `SDKP_PointingRowStep`

After installation, read back all 27 definitions and compare their exact
formula text with the prototype before connecting any worksheet formula.

### 2. Connect only `_Engine Master`

Apply these changes and no other engine-sheet change:

- Add the diagnostic label `_Engine Master!L101 = "Prototype Candidate Chain"`.
- Add the sole chain anchor
  `_Engine Master!L102 = SDKP_Chain($B$3:$J$11,8)`. Its spill must occupy
  `L102:T111` without `#SPILL!` or other formula error.
- Replace exactly `B15:J23` (81 formulas) with the prototype projection pattern:
  `=INDEX($L$102:$T$110,ROW()-14,COLUMN()-1)`.
- Replace `_Engine Master!B88` with the prototype chain-aware Coach formula.
  It must suppress guidance during Verify mode, conflicts, and dead-end state;
  return `N111` for `ASSIGNMENT`; report the executed count for `LIMIT`; report
  contradiction; and otherwise call `SDK_FallbackReveal`.
- Replace `_Engine Master!B89:B90` with the prototype formulas that split
  `B88`. Do not call `SDK_PlayerTrace` independently from these two cells.

Package comparison against production reports exactly **85 changed formula
cells** on `_Engine Master`: 81 in `B15:J23`, three in `B88:B90`, and the
single anchor in `L102`. `L101` is the only new constant in that engine area;
`L103:T111` are spill results, not individually written values or formulas.

Do not modify `_Engine`, `_Engine Medium`, `_Engine Hard`, or `_Engine Expert`.
They must retain the production used range through row 100 and their existing
`SDK_Candidates`, `SDK_PlayerHint`, and `SDK_PlayerTrace` formulas.

### 3. Bring the native pilot tests with the feature

- Update `_Tests!A145:B145` and `_Tests!D145` from the prototype so the
  player-safe Coach-path assertion recognizes the Master chain/fallback path.
  Preserve the existing expected/result cells in `C145` and `E145`.
- Copy the verified prototype block `_Tests!A164:E202` exactly, preserving its
  nearby format. This includes the structured action contracts, progress and
  contradiction guards, all seven strategy variants, chain stop states,
  valid-entry replay, wrong-entry recovery, and the eight-step boundary.
- The final boundary case at row 202 constructs nine independent Pointing-row
  patterns. With limit 8 it must return `LIMIT`, report eight executed steps,
  preserve the ninth pattern as the pending hint, and retain an eight-entry log.
  The assertion was mutation-checked: changing only its call limit from 8 to 7
  produced `FAIL`; restoring 8 produced `PASS`.

Current package comparison reports **95 formula differences** on `_Tests`
between production and the saved prototype. Treat `A164:E202` as one verified
test block rather than recreating formulas manually.

## Required invariants

Before and after the pilot, confirm all of the following:

- `01 Easy`, `02 Medium`, `03 Hard`, and `04 Expert` user and engine formulas
  are byte/formula equivalent to the source production copy.
- Master starts with 23 givens, temporary input cells
  `05 Master!C8,G5,K8,K17,K23,K26` blank, and `05 Master!AS11 = "Off"`.
- Workbook calculation mode remains `Automatic`; iteration and other workbook
  settings are not changed.
- No sheet is added, deleted, renamed, reordered, hidden, or unhidden.
- Existing names, validations, conditional formats, protection, views, and
  navigation are preserved.
- The original production workbook is never opened for write and retains its
  recorded SHA-256.

## Native acceptance sequence

1. Verify the fresh pilot copy matches the production hash before opening it.
2. Record the production-copy baseline before any change.
3. Install and read back the 27 names while the chain anchor is still absent.
4. Add the Master anchor, projection, Coach formulas, and test block in bounded
   phases, reading each affected range after its write.
5. Confirm `_Tests!E1:E202` is **131 PASS, 0 FAIL, 18 SKIP** and row 202 is
   `PASS`.
6. Confirm the clean Master fixture reports `ASSIGNMENT`, step `0`, with the
   Hidden Single at `R8C3 = 1`; confirm `B15:J23` equals `L102:T110`.
7. Repeat one normal correct-entry/delete round trip and one wrong-entry/clear
   round trip. Recheck candidate matrix, conflict count, chain state, Coach
   suppression, and exact recovery.
8. Search calculated values for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`,
   `#NUM!`, `#NULL!`, `#SPILL!`, and `#CALC!`. Do not treat the intentional
   formula-reference text examples as calculated errors.
9. Save through Excel, close, reopen the same pilot copy, rediscover the session,
   and repeat steps 5, 6, and 8. Run offline ZIP integrity and cached-result
   checks after close.
10. Compare all non-pilot sheets and all pre-existing defined names against the
    original production copy. Any unexplained difference fails the pilot.

Do not add repeated full-recalculation or memory stress loops to this acceptance
sequence. Stop only for normal-use delay or resource use that directly blocks
play.

## Stop and cut conditions

Discard the pilot copy and leave production unchanged if any of these occurs:

- recursive `#NUM!`, spill, name, or other calculated formula error;
- a save/reopen result differs from the pre-save result;
- input deletion or conflict clearing fails to restore the clean state;
- another engine or user sheet changes outside the exact pilot list;
- normal interaction becomes unusable or directly exceeds practical resources.

Do not respond to a failed pilot by building a fixed-depth clone. Retain the
existing single-step `SDK_PlayerHint`/`SDK_PlayerTrace` path and cut the
multi-step chain from the production proposal.

## Promotion sequence after a successful pilot

A successful Master production-copy pilot is evidence to request a separate
production rollout decision, not permission to edit production. If approved,
migrate one engine at a time in this order: Master, Expert, Hard, Medium, Easy.
After each engine, repeat its native candidate/Coach checks before proceeding.
All five production sheets must eventually use the same retained contract, or
the feature should remain prototype-only.
