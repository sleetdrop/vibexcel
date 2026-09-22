# Sudoku native Excel prototype: continuation checkpoint

## User-authorized task

Build a single-page candidate-elimination prototype in an independent test copy,
then measure native Microsoft Excel recalculation time and memory behavior for
input, deletion, and repeated recalculation. Keep the committed production
workbook unchanged. Close Excel after testing. This is prototype verification,
not authorization to deploy to all five production sheets or push Git.

## Current state

- Repository: `/Users/jiangyuan/Documents/learn/excel/vibexcel`
- Production file: `Sudoku/sudoku.xlsx`
- Last approved commit: `ef5d250 Fix Sudoku coach output and hint markers`
- Production SHA-256: `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`
- Last saved Excel regression results: 93 PASS, 0 FAIL, 18 SKIP.
- All five pages: mode Off, no player entries.
- Before preserving these notes, the Git worktree was clean and the production
  workbook still matched the recorded SHA-256.
- A native single-page prototype has now been created and tested on 2026-09-10.
  The preserved workbook is
  `Sudoku/dev-notes/candidate-chain-prototype/sudoku-candidate-chain-prototype.xlsx`.
  It is an independent test artifact, not the production workbook.
- The 2026-09-08 retry opened an isolated copy in Microsoft Excel and confirmed
  the ChatGPT pane was signed in. The connected-document tool descriptions were
  advertised, but `list_document_sessions` failed before sending a request with
  `TypeError: tools.mcp__codex_apps__codex_document_control_list_document_sessions
  is not a function`. The executable binding was absent even though the
  Spreadsheets plugin was enabled. Treat this as a Codex/plugin runtime blocker,
  not a workbook-registration failure.
- Excel was closed after the failed retry. The production workbook was not
  opened or changed in that retry.
- Two temporary workbook copies contained no intentional prototype formulas.
  They were not preserved because a fresh byte-for-byte copy can be recreated
  from `Sudoku/sudoku.xlsx`; one copy had only been re-saved by Excel.
- Prior Excel automation once reached 14.6 GB RAM. Keep live commands small;
  establish comparable memory baselines, record the measurement definition,
  and stop if memory grows substantially.

## 2026-09-10 native prototype result

- The Codex document-control bindings recovered after the app update. A fresh
  Excel session exposed read/write, workbook metadata, range image, and
  `run_officejs` tools for the exact test copy.
- The pre-fix failure was reproduced with the Master fixture: the coach said to
  remove 7 from R3C8, but the displayed candidate remained `2347` because every
  strategy regenerated candidates from the board.
- The test copy now contains workbook-level `SDKP_` LAMBDA names. They rebuild a
  bounded candidate-elimination chain from the current board, stop before a
  placement, and return one shared 9x9 candidate matrix plus status, step count,
  next hint, pending elimination, and chain log.
- Only `_Engine Master` is wired to the prototype. Its candidate display reads
  the shared spill at `L102:T111`; Hint/Trace read the same result. B89/B90 now
  parse B88 instead of independently recomputing the player trace.
- The reachable Master fixture passed exactly: Pointing row removed 7 from
  R3C8 (`2347 -> 234`), Naked pair row reduced R2C7 (`167 -> 67`), then the
  chain stopped at `Hidden single in column · R6C7 = 1` after two eliminations.
- Deleting K26 rebuilt from the changed board and restored R3C8 to `2347` and
  R2C7 to `167`; restoring K26 replayed the two-step chain. No exclusion history
  is stored.
- Conflict, stable/exhausted, and explicit budget exits passed. With budget 1,
  the status was `LIMIT` and the user-facing hint explicitly reported the
  limit instead of claiming that no strategy remained.
- Candidate-aware Excel formulas for all seven variants passed synthetic native
  tests: Pointing row/column, Claiming row/column, and Naked pair row/column/box.
- Three unchanged full recalculations measured 513 ms, 420 ms, and 356 ms
  inside the Office.js call (calculation plus `ctx.sync`, excluding connector
  round-trip latency).
- Excel RSS was sampled with `ps`: about 629,200 KiB after opening, 362,416 KiB
  after installing the prototype, and 1,355,216 KiB after the full logic suite
  plus three explicit full recalculations. Because that increase was substantial,
  the planned three extra input/deletion stress cycles were not run. The shorter
  deletion/replay behavior test had already passed. This does not establish a
  leak, but it is a rollout blocker until a longer matched-memory test is safe.
- The test workbook was saved, closed, and copied into the project. Excel was
  then quit and its process was confirmed absent. The production workbook was
  never opened in this round and retained SHA-256
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.

## 2026-09-10 matched memory diagnosis

- Excel was launched fresh and only the saved project prototype was opened. The
  production workbook was not opened. Main-process RSS was measured in KiB from
  the exact `/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel`
  process after each bounded phase.
- Setup/read-only samples were noisy and generally fell as Excel settled:
  486,752 at the start screen, 594,976 after opening the workbook, 510,528 after
  connecting the ChatGPT add-in, 435,584 after session discovery/metadata, and
  270,304 after the first bounded range read.
- One automatic input cycle did not reproduce the large spike. Clearing K26 and
  reading recalculated outputs used 298,032 KiB; restoring K26 and re-reading
  used 285,936 KiB. The candidate display correctly changed from `234`/`67` to
  `2347`/`167` on deletion and replayed the two-step chain on restoration.
- The first explicit full calculation took 495 ms and immediately raised RSS to
  961,504 KiB. After 30 seconds idle it fell to 656,720 KiB, still 370,784 KiB
  above the restored-input sample.
- The second explicit full calculation took 467 ms and immediately raised RSS
  to 1,172,592 KiB. After another 30 seconds idle it fell to 722,656 KiB. The
  second settled sample was 65,936 KiB above the first settled sample and
  436,720 KiB above the pre-full-calculation sample.
- Both calculations returned the same correct status: `ASSIGNMENT`, two
  elimination steps, then `Hidden single in column · R6C7 = 1`. This separates
  formula-result correctness from the memory problem: explicit workbook-wide
  recalculation produces a large temporary allocation and leaves some retained
  growth. Two samples do not prove an unbounded leak, but the behavior is large
  enough to keep production rollout blocked. Further repeated recalculations and
  malformed-candidate stress tests were deliberately stopped.
- The fixture was restored and verified before saving. A workbook-wide error
  search found no calculated formula errors; the only matches were three
  intentional `#N/A` text examples on `Formula Reference`. The saved prototype
  SHA-256 is now
  `ba76272e9a6a0e9aa81a908c97b1a5689b6dde6a7a234723eb9af8bc766fffac`.
  Excel was closed and quit, and its main process was confirmed absent. The
  production workbook remains byte-for-byte unchanged at the recorded hash.

## 2026-09-10 formula surface and A/B diagnosis

- The saved prototype contains exactly 17 workbook-level `SDKP_` names with
  17,714 formula characters. Only one worksheet formula calls that surface:
  `_Engine Master!L102 = SDKP_Chain($B$3:$J$11,8)`.
- `SDKP_ChainLoop` is recursive. At each visited state it references the full
  contradiction check, the basic-hint dispatcher, the advanced-hint dispatcher,
  and (when continuing) an 81-cell candidate-matrix rewrite before the recursive
  call. The contradiction formula contains four 9x9 `MAKEARRAY` scans.
- `SDKP_AdvancedHint` references all seven advanced strategies. Each strategy
  formula contains an 81-element `REDUCE`; Pointing and Claiming also construct
  multiple dynamic arrays, while Naked Pair variants repeatedly build masks and
  filtered coordinate arrays. `SDKP_BasicHint` similarly references Naked and
  three Hidden Single scans. Formula inspection establishes a large allocation
  surface, but does not assume whether Excel evaluates unused `LET` bindings or
  skipped `IF` branches eagerly.
- A fresh byte-for-byte production copy was used as the no-`SDKP_` control.
  After the add-in connected and sheet metadata was read, RSS was 431,792 KiB.
  One explicit full calculation took 470 ms and raised RSS to 831,488 KiB.
  After 30 seconds idle it fell to 384,176 KiB, below the pre-calculation sample.
  Therefore a several-hundred-MiB temporary peak is part of full calculation for
  the existing workbook even without the candidate-chain prototype.
- A fresh prototype copy was then used for an intermediate control. The sole
  chain anchor at `_Engine Master!L102` was cleared and verified blank, while all
  17 `SDKP_` names remained. Pre-calculation RSS was 442,688 KiB. One full
  calculation took 421 ms and raised RSS to 1,103,840 KiB. After 30 seconds idle
  it fell to 408,672 KiB, again below the pre-calculation sample.
- This separates peak allocation from retained growth. The defined formula
  surface increases the temporary full-calculation peak even when its only sheet
  invocation is disconnected, but that intermediate control released the peak.
  In the connected prototype, the prior matched test retained 370,784 KiB after
  the first 30-second idle period and another 65,936 KiB after the second. The
  current evidence therefore associates retained growth with execution of the
  candidate chain, while ordinary full calculation and the unused name surface
  account for much of the temporary peak.
- This is still a root-cause family rather than proof of one faulty Excel
  function. The next minimum experiment should vary only the chain budget in
  fresh saved/reopened copies: `SDKP_Chain(...,0)`, then 1, then 2. Budget 0
  evaluates one state without recursion; budget 1 permits one recursive step;
  budget 2 reaches the current two-elimination assignment fixture. Quit Excel
  between variants and stop when retained growth reappears.
- Both A/B workbooks were temporary copies. Their changes were discarded, Excel
  was quit after each phase, and the production and preserved prototype files
  were not modified in these tests.

## 2026-09-10 budget-zero diagnosis

- A fresh temporary copy of the saved prototype was changed only at the sole
  chain anchor, from `SDKP_Chain($B$3:$J$11,8)` to
  `SDKP_Chain($B$3:$J$11,0)`. The formula and its `LIMIT`, step `0` result were
  verified, the copy was saved, Excel was fully quit, and the saved copy was
  reopened before measurement.
- After the add-in had connected and the cold-start workbook had settled,
  pre-calculation RSS was 346,944 KiB. One explicit full calculation took
  2,388 ms and returned `LIMIT`, step `0`. Immediate RSS was 1,166,688 KiB;
  after 30 seconds idle it was still 921,968 KiB. That is 575,024 KiB
  (about 562 MiB) above the settled pre-calculation sample.
- The stop condition was therefore reached at budget zero. No recursive call
  and no 81-cell candidate rewrite can occur in this variant, so recursion
  depth is not required for the retained-memory behavior. The remaining suspect
  surface is the work performed for one chain state: contradiction scanning,
  basic-hint dispatch, advanced-hint dispatch, and/or Excel's evaluation of the
  `LET` bindings and skipped branches that reference them.
- Budget 1 and budget 2 were deliberately not run because they cannot determine
  which single-state component is responsible and would only increase memory
  pressure. The next minimum experiment should isolate one-state components in
  fresh saved/reopened copies, starting with contradiction-only, basic-only, and
  advanced-only anchors, and stopping at the first component that reproduces
  retained growth.
- The temporary budget-zero copy was closed without saving post-calculation
  cache changes, Excel was quit, and its main process was confirmed absent. The
  production workbook and the preserved prototype were not modified.

## 2026-09-10 contradiction-isolation setup blocker

- The next phase was prepared as a fresh temporary prototype copy named
  `sudoku-candidate-chain-contradiction-only.xlsx`. Excel opened it normally,
  the ChatGPT add-in pane showed the signed-in composer, and the intended minimal
  anchor was confirmed from the stored name definitions as
  `=SDKP_Contradiction($B$3:$J$11,SDK_CandidateMatrix($B$3:$J$11))`.
- The current Codex turn did not expose the connected-document session commands:
  the previously working session-list call was absent from the callable tool
  surface. Local configuration still had
  `[plugins."spreadsheets@openai-primary-runtime"] enabled = true`, so this was
  treated as a turn/tool-loading blocker rather than an Excel, workbook, or
  formula failure.
- No diagnostic formula was written and no measurement was run. The temporary
  copy was closed, Excel was quit, and the production workbook and preserved
  prototype remained untouched. Resume by recreating a fresh copy, writing the
  contradiction-only anchor through a newly discovered connected Excel session,
  saving, fully restarting Excel, and then taking the same pre/immediate/30-second
  RSS samples used for the budget-zero phase.

## 2026-09-16 contradiction-isolation connection timeout

- The connected Excel tools were available again. A fresh copy of the preserved
  prototype was changed only at `_Engine Master!L102` to
  `=SDKP_Contradiction($B$3:$J$11,SDK_CandidateMatrix($B$3:$J$11))`.
  The live read returned `FALSE`, and the copy was saved through the connected
  Excel session before the cold-start phase.
- Excel was fully quit and the saved copy was reopened. The ChatGPT pane showed
  the signed-in composer and the workbook registered, but the first metadata
  read timed out. The pane then lost its chat contents; it was closed and opened
  once, and a new connected session registered successfully. A minimal read of
  only `L102` also timed out.
- At the second timeout Excel showed `Ready`; its main process used 0.0% CPU and
  542,512 KiB RSS. No explicit full calculation was run, so there is no valid
  pre/immediate/30-second memory result for the contradiction component.
- Closing the cold-opened copy prompted to save automatic calculation state, so
  only that temporary diagnostic copy was saved. Offline package inspection then
  confirmed that `L102` still contains the contradiction-only formula with a
  cached false value. Excel was quit, and the production workbook and preserved
  prototype remained byte-for-byte unchanged.
- A durable copy of this exact diagnostic state is now stored at
  `Sudoku/dev-notes/candidate-chain-prototype/experiments/sudoku-candidate-chain-contradiction-only.xlsx`
  so continuation does not depend on `/private/tmp`. Resume by making a fresh
  temporary copy from this file, fully restarting Excel, reopening ChatGPT, and
  retrying one minimal `L102` read. If that succeeds, take a settled RSS sample,
  run exactly one full calculation, and record immediate and 30-second RSS.

## 2026-09-16 contradiction-isolation retry

- A new temporary copy was made from the durable contradiction-only experiment,
  with Excel confirmed fully absent before opening it. The workbook title, grid,
  ChatGPT ribbon button, signed-in composer, and one exact connected session were
  all verified.
- This retry skipped metadata discovery and immediately issued the minimum read:
  only `_Engine Master!L102`, using its already verified sheet ID. That command
  again timed out at the document-command boundary.
- Excel still showed `Ready`; the main process used 0.0% CPU and 548,272 KiB RSS
  after 1 minute 21 seconds. This reproduces the connection timeout across a new
  Excel process, a new workbook copy, and a new connected session; it is not an
  active Excel calculation stall.
- No Office.js full calculation was attempted, so the contradiction-only memory
  question remains unanswered. Only the temporary copy's automatic calculation
  state was saved when Excel requested it during close. Excel was then quit.
- Do not repeat the same cold-start read loop in the next turn without an external
  change to the Codex/Excel connection runtime. The durable experiment and exact
  resume recipe above remain sufficient for the next meaningful retry.

## 2026-09-16 production-copy connection control

- To distinguish a workbook/formula problem from a connection-runtime problem,
  a fresh temporary control workbook was copied byte-for-byte from the verified
  production `Sudoku/sudoku.xlsx`. Its SHA-256 before opening was
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`,
  matching production exactly. The control therefore contained none of the
  candidate-chain prototype's `SDKP_` names or anchor formulas.
- Excel was confirmed fully absent before the control copy was opened. The exact
  workbook title and grid, ChatGPT ribbon button, signed-in composer, and one
  connected session were verified. The session advertised both
  `read_sheets_metadata@1` and `read_ranges@1`; their live schemas were fetched
  before execution.
- The smallest workbook-wide read, `read_sheets_metadata@1`, timed out with
  `command_timeout` / `Arc command timed out` (command
  `cmd_e_9cd4d26ac4d081919d191334ac9c99b7`). After confirming Excel still showed
  `Ready` and rediscovering the same connected session, the one allowed retry
  with the same idempotency key returned the same timed-out command.
- At the failure boundary, the Excel main process used 0.0% CPU and 390,656 KiB
  RSS. No Office.js command, workbook calculation, formula edit, or production
  write was performed. Excel requested a save while closing the temporary copy;
  only that temporary copy was saved, then Excel was quit.
- This control reproduces the connected-read timeout in a workbook with no
  candidate-chain prototype content. The timeout is therefore independent of
  the contradiction/basic/advanced isolation formulas and currently blocks the
  native memory A/B experiment at the connection-command boundary. Do not use
  this timeout as evidence for or against the formula algorithm or its memory
  behavior.
- A one-sheet blank `Book1` control was then created in a new Excel process with
  the same signed-in add-in and tool surface. `read_sheets_metadata@1` succeeded
  in 6.4 seconds (command `cmd_e_e005399645d08191b6463c3b17a15925`), returning
  the empty visible `Sheet1`. The connection runtime is therefore not globally
  broken for all workbooks.
- A second fresh byte-identical production copy bypassed metadata discovery and
  read only `_Engine Master!A1` through `read_ranges@1`, using the sheet ID
  previously verified for that engine. It timed out, and the one permitted
  same-command retry returned the same timed-out command
  `cmd_e_725a3b6f2120819182ae8aa24b063256`.
- To remove sheet-ID ambiguity, a minimal read-only Office.js script explicitly
  selected `README!A1`, loaded only `values,formulas,text`, and called one
  `ctx.sync()`. It also timed out (command
  `cmd_e_6da2e84c7a0c8191b4ddba12c87ecc0c`). Excel remained `Ready`, with 0.0%
  CPU and 524,960 KiB RSS. The script did not request calculation or mutate the
  workbook.
- The first control copy had been saved natively by Excel while closing, changing
  its temporary SHA-256 to
  `83180358f40bace5597643d8029656abed4d069e2da7063c0c5e0b3e58656ab6`.
  Reopening that Excel-normalized copy did not help: metadata still timed out
  (command `cmd_e_fee3f9def7608191b86afcc694c6d2dd`). This rules out a defect
  that Excel's ordinary save normalization repairs.
- Offline OOXML inspection found no abnormally large used range: user sheets top
  out at `B1:BI39`, engines at `A1:T100`, and tests at `A1:P162`. The workbook
  contains 13 worksheets, about 6,000 cell formulas, 51 existing `SDK_` defined
  names, and a 237,124-byte `calcChain.xml`. These are candidate isolation axes,
  not yet proven causes.
- Current boundary: blank workbook commands succeed, while all tested connected
  read paths on the production workbook and its derivatives time out at their
  first workbook synchronization. The trigger belongs to the existing Sudoku
  workbook content/structure, not to the new candidate-chain prototype. Do not
  proceed directly to synthetic blank-workbook combinations: follow the
  documentation-first diagnostic plan below.
- Resume the contradiction-only memory isolation only after the production-copy
  metadata control succeeds. Until then, connected commands cannot measure the
  formula variants independently.

## 2026-09-16 documentation-first diagnostic correction

- The installed Excel for Mac reports version `16.112.4` and build
  `16.112.26090911`.
- Microsoft documents request/response payload and range-size limits for Excel
  add-ins, and recommends loading only necessary properties while minimizing
  `context.sync()` calls. The failing controls already load one cell and use one
  sync, far below the documented 5,000,000-cell read limit. Do not attribute this
  failure to payload size or add further chunking without contrary evidence.
- Microsoft documents the calculation chain as dynamic metadata that Excel owns,
  reorders, and saves. Do not delete or hand-edit `calcChain.xml` as an ad hoc
  experiment.
- Microsoft's supported workbook-recovery flow is the next diagnostic path on a
  fresh temporary copy: try Excel's Open and Repair workflow, and separately try
  opening with calculation set to Manual. Changing the Excel application-wide
  calculation preference requires explicit approval and must be restored after
  the test.
- Microsoft documents runtime logging on Mac through the Excel
  `CEFRuntimeLoggingFile` preference, but says it primarily captures host-level
  manifest/loading diagnostics rather than JavaScript console output. Enable it
  temporarily only if the supported recovery tests do not distinguish the cause,
  and restore the preference afterward.
- OfficeDev issue `office-js#697` records a historical `context.sync()` read hang
  that Microsoft labeled as an Office JavaScript API product bug and later fixed.
  It shows that this failure class can be a platform regression, but it involved
  a different Excel version and much larger reads, so it is analogous evidence,
  not a diagnosis of this workbook.
- No public OpenAI documentation was found that defines the internal
  `Arc command timed out` condition or its timeout boundary. Treat that message as
  an undocumented wrapper failure unless vendor guidance becomes available.
- Only after the official recovery/logging paths have been exhausted should a
  minimal synthetic reproduction be considered for a Microsoft/OpenAI bug
  report. Do not use synthetic combinations as the primary diagnostic method.

## 2026-09-16 Open and Repair availability check

- A new byte-identical temporary copy of production was created at
  `/private/tmp/sudoku-open-repair-20260916.PgWsBE/sudoku-open-repair.xlsx`.
  Its SHA-256 before the check was
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`,
  matching production exactly.
- In Excel for Mac 16.112.4, the standard Open dialog was navigated to that
  exact directory and the workbook row was selected. The fully expanded dialog
  exposed only an ordinary `Open` button. Neither the visible UI nor its
  accessibility tree exposed an adjacent dropdown, secondary action, or
  `Open and Repair` command.
- The dialog was cancelled without opening the workbook. No calculation mode,
  application preference, workbook content, or protected artifact was changed.
  The Excel Quit command was then issued; the UI-control session reset while
  waiting for post-quit state, so this records the issued command rather than a
  separately observed process-state confirmation.
- Therefore the Microsoft support article's documented `Open and Repair` path
  is unavailable in the current Mac UI; this is not evidence that repair was
  attempted or failed. Do not substitute an ordinary open/save or hand-edit the
  OOXML package and label that as the supported recovery test.
- The remaining supported recovery diagnostic is opening a fresh temporary copy
  after setting Excel calculation to Manual. Because that setting is
  application-wide, obtain explicit approval immediately before changing it,
  record its prior value, and restore that value after the test even if the
  workbook read still times out.

## 2026-09-16 Manual-calculation diagnostic

- After explicit approval, a new byte-identical production copy was created at
  `/private/tmp/sudoku-manual-calc-20260916.i6Zf57/sudoku-manual-calc.xlsx`.
  Its initial SHA-256 was
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`,
  matching production exactly.
- Excel's application-level calculation mode was inspected before the change and
  was `Automatic`. It was temporarily changed to `Manual`, and the temporary
  copy was opened. The workbook grid showed `README`, Excel reported `Ready`,
  and the status bar exposed `Calculate` with the explanation that results may
  be out of date because automatic calculation was disabled. This confirms the
  copy actually opened under Manual calculation.
- The ChatGPT add-in was installed and its pane opened, but no connected Excel
  session registered. Discovery was attempted after opening, after explicitly
  activating a workbook cell, and once more after closing and reopening the
  pane. All three checks returned no connected sessions.
- Because no session registered, no workbook command reached Office.js and the
  planned minimal metadata read was never executed. This experiment is
  inconclusive about whether Manual calculation changes the existing first-sync
  timeout; do not record it as a passed or failed read.
- With separate action-time approval, the temporary workbook was closed with
  `Don't Save`. Its on-disk SHA-256 remained identical to production. No
  workbook calculation command, formula edit, production write, or save was
  performed.
- Excel calculation mode was then restored to `Automatic` and visibly verified
  in the Calculation preference pane. The automatically created blank `Book1`
  was closed without saving. The Excel Quit command was issued; the UI-control
  session again reset while attempting the post-quit inventory, so process
  absence was not independently observed.
- Current boundary: both supported recovery diagnostics have now been attempted
  as far as Excel for Mac and the connection runtime allowed. Open and Repair is
  unavailable in the current Mac UI, while the Manual-calculation test was
  blocked before its first workbook command by session registration. Runtime
  logging is the next documentation-backed diagnostic, but it targets add-in
  host/manifest/loading events rather than JavaScript console output. Enable it
  only as a temporary, separately approved application preference and restore
  that preference afterward.

## 2026-09-17 runtime-logging diagnostic

- Microsoft documents Excel for Mac runtime logging through the
  `com.microsoft.Excel` `CEFRuntimeLoggingFile` preference and requires an
  Office restart. Before this diagnostic the preference did not exist. It was
  temporarily set to `codex_sudoku_runtime_20260916.txt`, Excel was started
  cleanly, and the preference was deleted afterward. A final read confirmed it
  was restored to the original absent state.
- A new byte-identical production copy was created at
  `/private/tmp/sudoku-runtime-log-20260916.m7m6xl/sudoku-runtime-log.xlsx`.
  Before opening and after closing without saving, its SHA-256 was
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`,
  matching production exactly.
- Unlike the preceding automatic-calculation controls, the ChatGPT add-in
  registered a connected session for this copy while runtime logging was
  enabled. `read_sheets_metadata@1` succeeded in 7.3 seconds and returned all
  13 worksheets (command `cmd_e_a9a60f79257881919b62b5229aff5346`).
- A one-cell `read_ranges@1` request for `01 Easy!K5` then succeeded in 5.1
  seconds (command `cmd_e_ecf4c776ca048191bfea2fbe158975e9`). Its response
  identified the requested sheet but emitted no value, formula, or text payload.
  This establishes that the command boundary completed, but it is not evidence
  that K5 has no formula and does not resolve the original display question.
- The 22-line runtime log records successful parsing and registration of the
  Excel Labs `LABS.GENERATIVEAI` custom function for `PERSONAL.XLSB` and the
  temporary workbook. It also records one `Unexpected Manifest` entry for
  `SolutionId:wa200010001`, version `1.0.0.1`, store locale `en-US`, reporting
  missing `DisplayName` and `Description`, followed by manifest-parsing starts.
  The log does not identify that solution as ChatGPT, and no supported source
  found during this round established its identity. Because session discovery
  and both read commands succeeded afterward, this entry is not established as
  causal or blocking.
- The final raw log is preserved at
  `diagnostics/runtime-logging-20260917/excel-runtime.log`; its SHA-256 is
  `b764b124fe9f9c4350b9ab8f9ab4bdfc845336374a7ab3618f340b6361ee4c4c`.
  Runtime logging is documented as host/manifest/loading diagnostics, not
  JavaScript console output, so the absence of a command-time error in this log
  does not prove that the connector path was healthy internally.
- With explicit action-time approval, the temporary workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. Production,
  the preserved prototype, and the durable contradiction-only experiment were
  not opened or changed in this diagnostic.
- Current inference: this clean logged run recovered the previously failing
  command path, but a single success cannot distinguish a transient runtime
  recovery from an effect of enabling logging. The next minimum diagnostic is a
  matched A/B rerun with logging disabled, using a new byte-identical production
  copy after a clean Excel restart. Resume contradiction-only memory isolation
  only after that no-logging production-copy metadata control succeeds.

## 2026-09-17 no-logging matched control

- Excel was confirmed absent before the test. The
  `com.microsoft.Excel` `CEFRuntimeLoggingFile` preference was also confirmed
  absent, so this run did not use runtime logging.
- A new byte-identical production copy was created at
  `/private/tmp/sudoku-no-log-control-20260917.f53ADD/sudoku-no-log-control.xlsx`.
  Its initial SHA-256 was
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`,
  matching production exactly.
- After a clean Excel start, the exact workbook title, worksheet grid, installed
  ChatGPT ribbon button, signed-in composer, and one connected session were
  verified. No other workbook session was selected.
- `read_sheets_metadata@1` succeeded and returned all 13 worksheets (command
  `cmd_e_a938b3409664819197857adf59c1ffba`; observed connector round trip about
  8.0 seconds). This is the matched production-copy control that had repeatedly
  timed out on 2026-09-16.
- A follow-up `read_ranges@1` request for `01 Easy!K5` also succeeded (command
  `cmd_e_53e3dc5bf5648191af32353b009c4d0a`; observed connector round trip about
  6.2 seconds). As in the logging-enabled run, the response identified the sheet
  but contained no value, formula, or text payload. This repeated response shape
  is independent of runtime logging; it still does not establish whether K5
  contains a formula or explain the original formula-bar display.
- With explicit action-time approval, the temporary workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The final
  temporary-file SHA-256 still matched production, and a direct byte comparison
  returned equality. The production, preserved prototype, and durable
  contradiction-only experiment retained their recorded hashes.
- Result: runtime logging is not required for the command path to succeed. The
  earlier timeouts were therefore transient or dependent on some other session
  state; this A/B pair does not identify that state. Do not treat the manifest
  warning from the logged run as the cause of the earlier timeouts.
- The prerequisite for resuming formula isolation is now satisfied: a fresh
  production-copy metadata control succeeds. The next minimum experiment can
  return to the durable contradiction-only workbook, using a fresh copy and
  clean Excel restart. First retry only `_Engine Master!L102`; if that succeeds,
  record settled RSS, run exactly one full calculation, and record immediate and
  30-second RSS. Stop if the minimal read times out again or memory rises sharply.

## 2026-09-17 contradiction-memory retry blocked by blank add-in pane

- Excel was confirmed absent before the test. Runtime logging remained disabled.
  A new temporary copy was created at
  `/private/tmp/sudoku-contradiction-memory-20260917.eHiUFH/sudoku-contradiction-memory.xlsx`
  from the durable contradiction-only experiment. Both files initially had
  SHA-256
  `e9af371a8fff9f62a5dc0fd0944f344658b87407da2a885fc0030279d0b2e945`.
- The exact workbook title and worksheet grid opened normally, Excel showed
  `Ready`, and the ChatGPT ribbon button was present. The ChatGPT pane initially
  opened blank. It was closed and reopened once, following the live-control
  recovery rule. The second load showed `Loading...`, then reached its hosted
  content URL, but the pane remained visually blank and exposed no `New chat`
  heading or `Ask anything` composer after two additional 10-second checks.
- Because the signed-in composer could not be verified, connected-session
  discovery was not attempted. No workbook command reached the document
  boundary, `_Engine Master!L102` was not read, and no Office.js calculation was
  run. This round therefore adds no evidence about contradiction-only formula
  memory behavior.
- At the stop boundary, Excel still reported `Ready`; its main process used 0.0%
  CPU and 489,328 KiB RSS after 3 minutes 6 seconds. This is evidence of an
  add-in-pane loading failure, not an active workbook-calculation stall.
- With explicit action-time approval, the temporary workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The final
  temporary file remained byte-identical to the durable experiment, and the
  production, preserved prototype, and durable experiment retained their
  recorded hashes.
- Do not bypass the setup gates or send commands to an unverified session. The
  next meaningful retry requires a fresh Excel process in which the ChatGPT pane
  visibly reaches the signed-in composer. Then repeat only the `L102` read before
  deciding whether to perform the one-calculation memory measurement.

## 2026-09-17 contradiction-only memory isolation completed

- Excel was confirmed absent before the test and runtime logging remained
  disabled. A fresh temporary copy was created at
  `/private/tmp/sudoku-contradiction-memory2-20260917.6YGXQz/sudoku-contradiction-memory2.xlsx`
  from the durable contradiction-only experiment. Both files had SHA-256
  `e9af371a8fff9f62a5dc0fd0944f344658b87407da2a885fc0030279d0b2e945`.
- After a clean Excel start, the ChatGPT pane visibly reached its signed-in
  composer and exactly one connected session was registered for the temporary
  workbook (`d2756971-ce80-4bd8-b71b-58d5c4d6254d`).
- The pre-calculation read of `_Engine Master!L102` succeeded (command
  `cmd_e_1dc0dde9bb448191adb78577d6fa3e8d`). It returned `FALSE` and formula
  `=SDKP_Contradiction($B$3:$J$11,SDK_CandidateMatrix($B$3:$J$11))`.
- Pre-calculation main-process RSS samples were 514,592 KiB, 495,984 KiB, and
  429,856 KiB. The last value is the comparison baseline, but RSS was still
  declining and CPU was 0.7%, so it is only a settled-ish baseline rather than
  a precise steady-state measurement.
- Exactly one explicit full-workbook calculation succeeded (command
  `cmd_e_0821bbffa6008191885bad3afc9d25ed`) and took 477 ms at the Office.js
  command boundary.
- The first post-calculation process sample was 782,560 KiB, 352,704 KiB
  (about 344.4 MiB) above the comparison baseline. The process-sampling command
  itself took 26.5 seconds, so this is not an immediate sample and the true peak
  may have been higher.
- A later sample was 468,688 KiB, only 38,832 KiB (about 37.9 MiB) above the
  comparison baseline. Because the requested wait plus sampling overhead placed
  it at least about 70 seconds after calculation completion, it must not be
  described as an exact 30-second sample.
- The post-calculation `L102` read also succeeded (command
  `cmd_e_194df9c07a088191a219522288a0b7c9`) and remained `FALSE` with the same
  formula.
- Result: the contradiction-only component did not reproduce the budget-zero
  chain's approximately 562 MiB retained growth. Most of its temporary
  allocation was released, so contradiction detection is unlikely to be the
  primary source of that retained-memory behavior. The timing limitations mean
  this result does not establish an exact peak or exact 30-second retention.
- With explicit action-time approval, the temporary workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The
  temporary copy remained byte-identical to the durable experiment; production
  and the preserved prototype retained their recorded hashes.
- Continue the component split with a fresh saved/reopened **basic-only** copy,
  clean Excel process, identical single-calculation protocol, and the same stop
  boundary. Test advanced-only only if basic-only does not reproduce the large
  retained-memory growth.

## 2026-09-17 basic-only memory isolation reproduced retained growth

- A fresh temporary copy of the preserved prototype was changed only at
  `_Engine Master!L102`, from `=SDKP_Chain($B$3:$J$11,8)` to
  `=SDKP_BasicHint(SDK_CandidateMatrix($B$3:$J$11))`. The write and readback
  succeeded through the exact connected workbook session (write command
  `cmd_e_eacd053c18408191b9947f1bb2e674aa`; verification command
  `cmd_e_1ab9e32c0c6881918ab349092923557d`).
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed that `L102` contains only the basic-hint anchor. This
  saved state is preserved at
  `experiments/sudoku-candidate-chain-basic-only.xlsx`, with SHA-256
  `7bdd7bd1869957b59ae8654e236d576d03b5c48971ed938fcff3ac041f9f8704`.
- The memory phase used a new byte-identical copy at
  `/private/tmp/sudoku-basic-only-memory-20260917.vDobTe/sudoku-basic-only-memory.xlsx`
  after another clean Excel start. The ChatGPT pane visibly reached its signed-in
  composer and the exact measurement workbook registered as a new connected
  session (`bp_arc_e_6aab86089c408191944b22094faf15d6`).
- The pre-calculation `L102` read succeeded (command
  `cmd_e_729a923da1c08191be8dffe2dc663c25`). It returned the expected formula and
  no value or text payload, consistent with the starting state having no
  immediate basic hint.
- Pre-calculation main-process RSS samples were 471,920, 437,488, 426,016,
  425,984, 419,136, and 420,512 KiB. The comparison baseline is the median of
  the final three samples, 420,512 KiB. Those samples still ranged from 0.0% to
  0.6% CPU, so this is a settled-ish comparison baseline rather than an exact
  steady-state measurement.
- Exactly one explicit full-workbook calculation succeeded (command
  `cmd_e_8d2bd52ccc848191967418b6fe7c9789`) and took 531 ms at the Office.js
  command boundary.
- The first available post-calculation sample was 918,864 KiB, 498,352 KiB
  (about 486.7 MiB) above the comparison baseline. The process-sampling command
  completed about five seconds after the calculation command returned, so this
  is not an immediate peak and the true peak may have been higher.
- A later sample was 824,960 KiB, still 404,448 KiB (about 395.0 MiB) above the
  comparison baseline. It was taken 43 seconds after the first post-calculation
  sample and roughly 48 seconds after the calculation command returned, not at
  an exact 30-second boundary.
- The post-calculation `L102` read succeeded (command
  `cmd_e_2912581c9d28819184a74a590b16209c`) and retained the same formula with
  no value or text payload.
- Result: basic-only reproduced substantial retained-memory growth after one
  full calculation. The retained delta was smaller than the budget-zero
  chain's approximately 562 MiB result, but large enough to satisfy the stop
  condition. No second calculation and no advanced-only phase were run.
- This narrows the current root-cause family to the work referenced by
  `SDKP_BasicHint`: `SDKP_Naked`, `SDKP_HiddenRow`, `SDKP_HiddenCol`, and
  `SDKP_HiddenBox`, including possible evaluation of their `LET` bindings. It
  does not yet identify one faulty function or prove Excel's branch-evaluation
  behavior.
- With explicit action-time approval, the measurement workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The setup
  copy, durable basic-only experiment, and measurement copy remained
  byte-identical at the recorded hash. Production, the preserved prototype, and
  the contradiction-only experiment retained their recorded hashes.
- Next isolate the four basic components in fresh saved/reopened copies, one at
  a time, starting with the lightweight `SDKP_Naked` anchor. If Naked stays
  bounded, continue with Hidden Row, Hidden Column, then Hidden Box, quitting
  Excel between variants and stopping at the first retained-growth reproduction.
  Do not run advanced-only while the basic branch remains the smaller confirmed
  root-cause family.

## 2026-09-17 Naked-only memory isolation reproduced retained growth

- A fresh temporary copy of the preserved prototype was changed only at
  `_Engine Master!L102`, from `=SDKP_Chain($B$3:$J$11,8)` to
  `=SDKP_Naked(SDK_CandidateMatrix($B$3:$J$11))`. The write and readback
  succeeded through the exact setup workbook session (write command
  `cmd_e_48fb9243b3cc8191ad92c23994ec6883`; verification command
  `cmd_e_0a22e40850cc8191b892a23e5650eb71`).
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed that `L102` contains only the Naked-only anchor. This
  saved state is preserved at
  `experiments/sudoku-candidate-chain-naked-only.xlsx`, with SHA-256
  `a1ed219adf0cfd5d448a6ac0bb069c2da7de24b8e80fbb1919b7ebfa776169c5`.
- The memory phase used a new byte-identical copy at
  `/private/tmp/sudoku-naked-only-memory-20260917.4ZRpu0/sudoku-naked-only-memory.xlsx`
  after another clean Excel start. The ChatGPT pane visibly reached its signed-in
  composer and the exact measurement workbook registered as a new connected
  session (`bp_arc_e_6aac0b63fe4c81918fe3a5253578bf53`).
- The pre-calculation `L102` read succeeded (command
  `cmd_e_abcf6d55b2208191aab921e4e735874b`). It returned the expected Naked-only
  formula and no value or text payload. The post-calculation read also succeeded
  (command `cmd_e_6857826aa60c8191846d60c0a8548c1d`) and retained the same formula
  with no value or text payload.
- Pre-calculation main-process RSS samples were 580,528, 578,800, 578,912,
  539,072, 562,688, 568,576, 559,968, 531,776, 531,824, 531,824, and
  532,224 KiB. The comparison baseline is the median of the final three
  samples, 531,824 KiB. Those three samples were stable within 400 KiB and used
  0.0% to 0.1% CPU.
- Exactly one explicit full-workbook calculation succeeded (command
  `cmd_e_7445521595348191af7909711189c859`) and took 499 ms at the Office.js
  command boundary.
- The first available post-calculation sample was 1,347,632 KiB, 815,808 KiB
  (about 796.7 MiB) above the comparison baseline. It was taken about three to
  four seconds after the calculation command returned, not at the true peak.
- A later sample was 1,194,448 KiB, still 662,624 KiB (about 647.1 MiB) above
  the comparison baseline. It was taken 40 seconds after the first sample and
  roughly 43 to 44 seconds after command return, not at an exact 30-second
  boundary.
- Result: Naked-only reproduced substantial retained-memory growth after one
  full calculation, satisfying the stop condition. No second calculation,
  Hidden Single variant, or advanced-only phase was run. Although the later
  delta in this run was larger than the basic-only and budget-zero results,
  cross-run baselines and process noise do not establish that Naked intrinsically
  retains more memory than those variants.
- The current root-cause family is the `SDKP_Naked` path and/or the
  `SDK_CandidateMatrix` argument it evaluates. Contradiction-only also evaluated
  `SDK_CandidateMatrix` and stayed bounded, but it combined that input with a
  much heavier contradiction scan and returned a different scalar result. The
  evidence therefore points more strongly to the Naked path without yet
  separating Naked's own expression from candidate-matrix materialization.
- With explicit action-time approval, the measurement workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The setup
  copy, durable Naked-only experiment, and measurement copy remained
  byte-identical at the recorded hash. Production, the preserved prototype, and
  the contradiction-only and basic-only experiments retained their recorded
  hashes.
- The next minimum split should use a fresh saved/reopened
  `=SDK_CandidateMatrix($B$3:$J$11)` anchor. Its 9x9 spill fits within the
  prototype's existing `L102:T110` output area. If matrix-only stays bounded,
  add Naked's expression stages one at a time: `LEN(matrix)=1`, then `TOCOL`,
  then `XMATCH`, stopping at the first retained-growth reproduction. If
  matrix-only itself reproduces the behavior, investigate `SDK_CandidateMatrix`
  before returning to Naked. Do not run Hidden Single variants or advanced-only
  while this smaller split remains available.

## 2026-09-18 matrix-only setup blocked after add-in update

- Excel was confirmed absent, runtime logging remained disabled, and a fresh
  setup copy was created from the preserved prototype before opening Excel.
  Production, the preserved prototype, and the Naked-only experiment all
  matched their recorded hashes.
- Opening ChatGPT for Excel displayed a mandatory `Update Add-in` screen. With
  explicit approval, `Update now` was selected. The workbook had not been
  edited. Excel nevertheless marked the temporary copy as changed, so the copy
  was closed with `Don't Save` after separate explicit approval.
- Excel was fully restarted and a second fresh setup copy was created from the
  preserved prototype at
  `/private/tmp/sudoku-matrix-only-setup2-20260918.z2epxR/sudoku-matrix-only-setup2.xlsx`.
  The updated ChatGPT pane opened, but its content remained visually blank and
  exposed neither `New chat` nor the `Ask anything` composer.
- The pane was closed and reopened once, following the live-control recovery
  rule. It remained blank. Because sign-in could not be verified, connected
  session discovery was not attempted. No workbook command was sent, no formula
  was written, and no calculation or memory measurement was run.
- With explicit approval, the second untouched setup copy was closed with
  `Don't Save`. Excel was quit and its process was confirmed absent. Both setup
  copies remained byte-identical to the preserved prototype with SHA-256
  `ba76272e9a6a0e9aa81a908c97b1a5689b6dde6a7a234723eb9af8bc766fffac`.
  Production remained unchanged at its recorded hash.
- The matrix-only experiment remains the next diagnostic step. Do not bypass
  the sign-in gate or send commands to an unverified session. Resume only after
  a clean Excel start in which the ChatGPT pane visibly reaches the signed-in
  composer; then create another fresh copy and continue with only the
  `=SDK_CandidateMatrix($B$3:$J$11)` anchor.

## 2026-09-18 matrix-only memory isolation reproduced retained growth

- A later clean retry resolved the updated add-in's blank-pane state without
  another update or workaround. The pane reached `New chat` and the signed-in
  composer, and the exact setup workbook registered as connected session
  `bp_arc_e_6aac159313308191aed017db0823b9b9`.
- A fresh prototype copy was changed only at `_Engine Master!L102`, from
  `=SDKP_Chain($B$3:$J$11,8)` to
  `=SDK_CandidateMatrix($B$3:$J$11)`. The original read, write, and 9x9 spill
  verification succeeded (commands `cmd_e_155d3a73ade88191a834833a84e87740`,
  `cmd_e_4d2920a8f7288191a4a9d67f787be2ef`, and
  `cmd_e_2efe1362d7988191b487f6d4749c8c53`).
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed the array anchor and spill reference `L102:T110`. The
  saved state is preserved at
  `experiments/sudoku-candidate-chain-matrix-only.xlsx`, with SHA-256
  `610ba9ba71e9340a189b9af1f5c96c07b72e4aee919574dd7e718d5ad74c14d4`.
- The memory phase used a new byte-identical copy at
  `/private/tmp/sudoku-matrix-only-memory-20260918.b2OfQh/sudoku-matrix-only-memory.xlsx`
  after another clean Excel start. It registered as connected session
  `bp_arc_e_6aac163e050c8191ae6a47fba9b81ff8`.
- The pre-calculation read of `L102:T110` succeeded (command
  `cmd_e_18eb46e8c1788191b79b92abd430a731`) and returned the expected formula
  and candidate matrix. Pre-calculation main-process RSS samples were 735,088,
  733,744, 733,872, 733,808, 733,888, and 733,824 KiB. The comparison baseline
  is the median of the final three samples, 733,824 KiB; those samples differed
  by only 80 KiB and used 0.1% to 0.3% CPU.
- Exactly one explicit full-workbook calculation succeeded (command
  `cmd_e_ed77647040208191b192f7f7136d5ff7`) and took 457 ms at the Office.js
  command boundary.
- The first available post-calculation sample was 1,338,112 KiB, 604,288 KiB
  (about 590.1 MiB) above the baseline. It was taken about three seconds after
  the calculation command returned, not at the true peak.
- A later sample was 1,336,336 KiB, still 602,512 KiB (about 588.4 MiB) above
  the baseline. It was taken 42 seconds after the first sample and roughly
  45 seconds after command return, not at an exact 30-second boundary.
- The post-calculation read succeeded (command
  `cmd_e_53254f12a030819189f7ba3293906fca`) and returned the same formula and
  candidate matrix. Result: matrix-only reproduced substantial retained-memory
  growth and released almost none of the first observed delta during the later
  interval. The stop condition was satisfied, so no second calculation and no
  `LEN`, `TOCOL`, or `XMATCH` phase was run.
- This removes `SDKP_Naked` from the minimum reproducer. The remaining surface
  is direct worksheet evaluation of `SDK_CandidateMatrix`, whose definition is
  `LAMBDA(board,MAKEARRAY(9,9,LAMBDA(r,c,SDK_Candidates(board,r,c))))`, and/or
  materialization of its 9x9 spill. Contradiction-only also receives the matrix
  as an argument but returns a scalar and stayed bounded, so the evidence does
  not yet distinguish candidate generation from direct spill materialization.
- With explicit action-time approval, the measurement workbook was closed using
  `Don't Save`. Excel was quit and its process was confirmed absent. The setup
  copy, durable experiment, and measurement copy remained byte-identical at the
  recorded hash. Production and the preserved prototype retained their recorded
  hashes.
- The next two orthogonal controls should be run in fresh saved/reopened copies,
  one per clean Excel process. First isolate the array scaffold and spill with
  `=MAKEARRAY(9,9,LAMBDA(r,c,r*10+c))`. If that stays bounded, isolate one
  representative candidate calculation without a spill using
  `=SDK_Candidates($B$3:$J$11,1,3)`, which should return `28` for the current
  fixture. If both stay bounded, increase candidate-call scale through 1x9 and
  3x9 `MAKEARRAY` anchors before returning to 9x9. Stop at the first retained-
  growth reproduction; do not resume Naked, Hidden Single, or advanced-only.

## 2026-09-18 constant MAKEARRAY reproduced retained growth

- A fresh copy of the preserved prototype was changed only at
  `_Engine Master!L102`, from `=SDKP_Chain($B$3:$J$11,8)` to
  `=MAKEARRAY(9,9,LAMBDA(r,c,r*10+c))`. The 9x9 spill was verified as rows
  `11..19` through `91..99`.
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed the array anchor and spill reference `L102:T110`. The
  saved state is preserved at
  `experiments/sudoku-candidate-chain-makearray-constant.xlsx`, with SHA-256
  `eb4aa7a8918a91ea8c8c377f3da259bfa2c1df15adf1563e73bd62a529504a67`.
- The memory phase used a new byte-identical copy at
  `/private/tmp/sudoku-makearray-constant-memory-20260918.rNAQDr/sudoku-makearray-constant-memory.xlsx`
  after another clean Excel start. It registered as connected session
  `bp_arc_e_6aacbac0abe88191ba1cb573baa9b44a`.
- The comparison baseline was 506,496 KiB, the median of the final three stable
  samples (506,496, 506,496, and 506,480 KiB; range 16 KiB). Exactly one full
  calculation succeeded (command `cmd_e_595a7f41d83081919b462d56bfa6b071`)
  and took 410 ms at the Office.js command boundary.
- The first available post-calculation sample was 1,330,000 KiB, 823,504 KiB
  (about 804.2 MiB) above baseline. A later sample was 1,340,832 KiB, still
  834,336 KiB (about 814.8 MiB) above baseline, about 41 seconds after the
  first sample. The formula and all spilled values remained unchanged.
- Result: a constant 9x9 `MAKEARRAY` spill reproduced substantial retained-
  memory growth without calling any Sudoku candidate function. Candidate
  generation is therefore no longer part of the minimum reproducer. The
  remaining surface is dynamic-array worksheet calculation/materialization,
  or its interaction with the workbook calculation graph.
- The stop condition was satisfied. No scalar candidate test, scaled candidate
  calls, second calculation, `LEN`, `TOCOL`, or `XMATCH` phase was run.
- After explicit action-time approval, the measurement workbook was closed
  with `Don't Save` and Excel was quit. The Excel process was confirmed absent.
  The measurement copy remained byte-identical to the durable setup workbook
  (`cmp=0`), and both retained SHA-256
  `eb4aa7a8918a91ea8c8c377f3da259bfa2c1df15adf1563e73bd62a529504a67`.
  Production and the preserved prototype also retained their recorded hashes.
  Full evidence is in
  `diagnostics/makearray-constant-memory-20260918/SUMMARY.md`.
- Next, use a fresh saved/reopened workbook to distinguish `MAKEARRAY` itself
  from multi-cell spill materialization, starting with a scalar 1x1
  `MAKEARRAY` or another orthogonal non-spilling control. Keep one variable per
  run and stop at the first retained-growth reproduction.

## 2026-09-18 1x1 MAKEARRAY also reproduced retained growth

- A fresh copy of the preserved prototype was changed only at
  `_Engine Master!L102`, from `=SDKP_Chain($B$3:$J$11,8)` to
  `=MAKEARRAY(1,1,LAMBDA(r,c,11))`. The result was `11`, and `L102:T110`
  contained no other value.
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed a dynamic-array formula whose reference is only
  `L102`. The saved state is preserved at
  `experiments/sudoku-candidate-chain-makearray-1x1.xlsx`, with SHA-256
  `1e3e7305525407d9acb96471fdf98da4b0ce193f327c2a8697587bf6cdaa20c4`.
- The memory phase used a new byte-identical copy at
  `/private/tmp/sudoku-makearray-1x1-memory-20260918.PhjoFQ/sudoku-makearray-1x1-memory.xlsx`
  after another clean Excel start. It registered as connected session
  `bp_arc_e_6aacbe0343bc8191ae271d0f7a0daaff`.
- The baseline required a longer settling period. Its final three samples were
  all 479,776 KiB. Exactly one full calculation succeeded (command
  `cmd_e_27eca35fd70c8191a033e921f3313ee2`) and took 380 ms at the Office.js
  command boundary.
- The first available post-calculation sample was 1,304,816 KiB, 825,040 KiB
  (about 805.7 MiB) above baseline. A later sample was 1,313,776 KiB, still
  834,000 KiB (about 814.5 MiB) above baseline, about 44 seconds after the
  first sample. The formula and value `11` remained unchanged.
- Result: a single-cell `MAKEARRAY` result reproduced substantial retained-
  memory growth. Multi-cell spill materialization and spill size are not
  necessary for reproduction. This still does not prove that `MAKEARRAY`
  itself is the trigger because the complete existing workbook was
  recalculated.
- The stop condition was satisfied. No second calculation or additional
  formula variant was run.
- After explicit action-time approval, the measurement workbook was closed
  with `Don't Save` and Excel was quit. The Excel process was confirmed absent.
  The measurement copy remained byte-identical to the durable setup workbook
  (`cmp=0`), and both retained SHA-256
  `1e3e7305525407d9acb96471fdf98da4b0ce193f327c2a8697587bf6cdaa20c4`.
  Production and the preserved prototype also retained their recorded hashes.
  Full evidence is in
  `diagnostics/makearray-1x1-memory-20260918/SUMMARY.md`.
- Next, repeat the same protocol with an ordinary scalar formula `=11`. If it
  stays bounded, the dynamic-array formula classification or `MAKEARRAY`
  evaluation is the minimum differentiator. If it also grows, investigate the
  anchor replacement and whole-workbook calculation separately.

## 2026-09-18 ordinary scalar formula also reproduced retained growth

- A fresh copy of the preserved prototype was changed only at
  `_Engine Master!L102`, from `=SDKP_Chain($B$3:$J$11,8)` to the ordinary
  scalar formula `=11`. Direct readback confirmed that `L102:T110` then
  contained only `L102 = 11`.
- The edited copy was saved, Excel was fully quit, and offline package
  inspection confirmed `<c r="L102"><f>11</f><v>11</v>` with neither array
  type nor spill reference. The clean saved state is preserved at
  `experiments/sudoku-candidate-chain-scalar-11.xlsx`, SHA-256
  `15cc34b9cfa646998271fbc4c289fa93b4657c407bbc17d23253d2442250fc5b`.
- The measurement phase used a fresh byte-identical copy at
  `/private/tmp/sudoku-scalar-11-memory-20260918.TfXoAP/sudoku-scalar-11-memory.xlsx`
  after another clean Excel start. It connected as session
  `bp_arc_e_6aacc95acb808191a28ef416c9cfc3ee`.
- The stable RSS samples were 609,520, 609,456, and 609,424 KiB; their median,
  609,456 KiB, is the baseline. Exactly one full calculation succeeded
  (command `cmd_e_1e3877635ed881919ecccb20337b10f9`) and took 444 ms at the
  Office.js command boundary.
- The first available post-calculation sample was 1,362,512 KiB, 753,056 KiB
  (about 735.4 MiB) above baseline. About 75 seconds later RSS was 1,343,824
  KiB, still 734,368 KiB (about 717.2 MiB) above baseline. The formula and
  value remained unchanged.
- Result: an ordinary scalar formula reproduced substantial retained-memory
  growth. A multi-cell spill, dynamic-array formula classification, and
  `MAKEARRAY` evaluation are therefore not required for reproduction. The
  saved anchor replacement and/or full-workbook recalculation of the remaining
  workbook graph is now the minimum unresolved surface. The earlier bounded
  contradiction-only run cannot be explained merely by its scalar output.
- The stop condition was satisfied; no second calculation was run. After
  explicit action-time approval, the measurement workbook was closed with
  `Don't Save` and Excel was quit. The process was confirmed absent. The
  measurement copy remained byte-identical to the durable setup (`cmp=0`) at
  the recorded hash; production and the preserved prototype also retained
  their recorded hashes. Full evidence is in
  `diagnostics/scalar-11-memory-20260918/SUMMARY.md`.
- Next, repeat the same protocol with a static numeric value `11` at `L102`
  (not a formula). If that stays bounded, the presence of any formula at the
  anchor is the differentiator. If it also grows, compare the saved anchor /
  calculation-chain state and a matched whole-workbook calculation control.

## 2026-09-20 static numeric value also reproduced retained growth

- A fresh copy of the preserved prototype was changed only at
  `_Engine Master!L102`, replacing `=SDKP_Chain($B$3:$J$11,8)` and its spill
  with the static numeric value `11`.
- The setup copy was saved and Excel was fully quit. Offline Open XML
  inspection confirmed `<c r="L102"><v>11</v></c>` with no formula node. The
  live range serializer's `<f>11</f>` output for this cell is therefore an
  interface artifact, not an Excel formula.
- The clean saved state is preserved at
  `experiments/sudoku-candidate-chain-static-11.xlsx`, SHA-256
  `65ee149f559643c3c6fb8112745214d7e160fe33f8b716e10f3447a95aed083f`.
- The measurement phase used a fresh byte-identical copy at
  `/private/tmp/sudoku-static-11-memory-20260920.wZMx3j/sudoku-static-11-memory.xlsx`
  after another clean Excel start. It connected as session
  `bp_arc_e_6aaffacf2ccc819193e64d28ec21ff39`.
- Settling RSS samples were 524,096, 498,656, 498,704, and 498,752 KiB. The
  median of the final three, 498,704 KiB, is the baseline. Exactly one full
  calculation succeeded (command `cmd_e_646434cf1d788191a6f4ab4160be908c`)
  and took 717 ms at the Office.js command boundary.
- The first available post-calculation sample was 913,712 KiB, 415,008 KiB
  (about 405.3 MiB) above baseline. After partial release, the final observed
  sample was 797,008 KiB, still 298,304 KiB (about 291.3 MiB) above baseline
  roughly 49 seconds after the first sample. The static value remained `11`.
- Result: a formula at `L102` is not required for reproduction. Formula
  content and formula classification at that anchor are therefore not the
  differentiator. Because this was still a full-workbook calculation, the
  result does not establish that the constant caused the growth; the remaining
  surface is the saved anchor replacement / calculation-chain state versus the
  existing workbook graph.
- The stop condition was satisfied and no second calculation was run. The
  Excel `CEFRuntimeLoggingFile` preference was absent. After explicit
  action-time approval, the measurement workbook was closed with `Don't Save`,
  Excel was quit, and its process was confirmed absent. The measurement copy
  remained byte-identical to the durable setup (`cmp=0`) at the recorded hash;
  production and the preserved prototype retained their recorded hashes. Full
  evidence is in
  `diagnostics/static-11-memory-20260920/SUMMARY.md`.
- Next, run the same protocol with `L102` blank. Then run a matched unmodified
  preserved-prototype full-calculation control. Together these controls should
  distinguish an anchor-replacement/calculation-chain effect from a general
  full-workbook calculation effect.

## 2026-09-20 blank L102 also reproduced retained growth

- A fresh copy of the preserved prototype was changed only by clearing the
  contents of `_Engine Master!L102`, removing the original
  `=SDKP_Chain($B$3:$J$11,8)` formula and its spill. Direct readback and a
  range image confirmed that `L102:T110` was empty.
- The setup copy was saved and Excel was fully quit. Offline Open XML
  inspection confirmed that `sheet13.xml` contains neither an `L102` cell
  node nor the removed target formula. The clean saved state is preserved at
  `experiments/sudoku-candidate-chain-blank-l102.xlsx`, SHA-256
  `1c817fed0d808ddd5e052badc2fd5b3c0f8cbba57420690d370487bf1565cb58`.
- The measurement phase used a fresh byte-identical copy at
  `/private/tmp/sudoku-blank-l102-memory-20260920.DJQ4s2/sudoku-blank-l102-memory.xlsx`
  after another clean Excel start. It connected as session
  `bp_arc_e_6aafffcf4cac81919254fdd095c15cea`.
- The final three settling RSS samples were 491,840, 491,872, and 491,792
  KiB, an 80 KiB range. Their median, 491,840 KiB, is the baseline. Exactly
  one full calculation succeeded (command
  `cmd_e_ea8ee216640081919f22b108cdfeb2f6`) and took 527 ms at the Office.js
  command boundary.
- The first available post-calculation sample was 1,245,984 KiB, 754,144 KiB
  (about 736.5 MiB) above baseline. The final observed sample, roughly 100
  seconds later, was 787,680 KiB, still 295,840 KiB (about 288.9 MiB) above
  baseline. `L102:T110` remained empty.
- Result: no formula, numeric value, or other content at `L102` is required
  for reproduction. The anchor contents are not the differentiator among the
  tested variants. This still leaves two possibilities: saving after removing
  or replacing the anchor changed relevant workbook state, or full calculation
  of the remaining workbook graph is sufficient by itself.
- The stop condition was satisfied and no second calculation was run. The
  Excel `CEFRuntimeLoggingFile` preference was absent. After explicit
  action-time approval, the measurement workbook was closed with `Don't Save`,
  Excel was quit, and its process was confirmed absent. The measurement copy
  remained byte-identical to the durable setup (`cmp=0`) at the recorded hash;
  production and the preserved prototype retained their recorded hashes. Full
  evidence is in
  `diagnostics/blank-l102-memory-20260920/SUMMARY.md`.
- Next, run a matched unmodified preserved-prototype control: copy the
  preserved prototype byte-for-byte, make no workbook edits, open it after a
  clean Excel start, and execute the same single full calculation. This is the
  decisive comparison between an anchor-edit/saved-state effect and a general
  full-workbook calculation effect.

## 2026-09-21 unmodified prototype also reproduced retained growth

- A fresh measurement copy was made directly from the preserved prototype.
  Before opening Excel, it was byte-identical (`cmp=0`) at SHA-256
  `ba76272e9a6a0e9aa81a908c97b1a5689b6dde6a7a234723eb9af8bc766fffac`.
  No workbook edit or save was performed before calculation.
- Only this workbook was open. Its original
  `_Engine Master!L102 = SDKP_Chain($B$3:$J$11,8)` formula and 9x9 spill were
  verified before and after calculation.
- The final stable RSS samples were 509,088, 509,152, and 509,264 KiB; their
  median, 509,152 KiB, is the baseline. Exactly one full calculation succeeded
  (command `cmd_e_562b36bf17b481919fbed79268189f00`) and took 488 ms at the
  Office.js command boundary.
- The first available post-calculation sample was 1,123,296 KiB, 614,144 KiB
  (about 599.8 MiB) above baseline. The final sample, roughly 91 seconds later,
  was 842,976 KiB, still 333,824 KiB (about 326.0 MiB) above baseline.
- Result: editing or saving `L102` is not required. The formula, static-value,
  and blank-anchor experiments were all confounded by the same whole-
  application full calculation, which is sufficient to reproduce the retained
  growth in the unmodified prototype.
- Microsoft documents `Application.calculate(full)` as recalculating all open
  workbooks after marking all cells dirty. Although only the target workbook
  was open here, this operation did not isolate a particular sheet or formula.
  Supported narrower APIs exist for calculating a worksheet or a range.
- Next, use a new byte-identical prototype copy in a clean Excel process and
  calculate only the `_Engine Master` used range (falling back to that worksheet
  if needed). Test other engine sheets only in separate clean processes and
  stop at the first reproducer. Do not continue changing `L102` content.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in `diagnostics/unmodified-prototype-memory-20260921/SUMMARY.md`.

## 2026-09-21 `_Engine Master` used range reproduced part of the growth

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy was not saved.
- The final three settling RSS samples were 379,472, 379,440, and 379,440 KiB;
  their median, 379,440 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on the sheet's used
  range, `_Engine Master!A1:T111` (command
  `cmd_e_78450924a0b88191b4b0c10d48cf0695`). It took 159 ms at the Office.js
  command boundary.
- The first post-calculation sample was 754,800 KiB, about 366.6 MiB above
  baseline. Roughly 142 seconds later RSS had fallen to 466,752 KiB but was
  still 87,312 KiB (about 85.3 MiB) above baseline.
- The original chain formula, candidate spill, `ASSIGNMENT` status, two-step
  count, hint, and elimination log were identical before and after calculation.
  No second calculation was run.
- Result: `_Engine Master` alone can produce a large temporary allocation and
  about 85 MiB of retained growth. This is materially smaller than the roughly
  326 MiB retained after application-level full calculation, so other sheets,
  cross-sheet combination, or full-dirty semantics account for the remainder.
- Next, in a separate clean process, calculate only `_Engine!A1:T100`, which
  backs `01 Easy` and has no prototype chain anchor. This is the most relevant
  existing-engine comparison. If it reproduces, continue one engine sheet per
  clean process; if it stays bounded, subdivide `_Engine Master`.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in `diagnostics/master-used-range-memory-20260921/SUMMARY.md`.

## 2026-09-21 `01 Easy` engine also reproduced retained growth

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy was not saved.
- The final three settling RSS samples were 398,704, 397,856, and 397,280 KiB;
  their median, 397,856 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on `_Engine!A1:T100`
  (command `cmd_e_fa13aebac08081919d1abb3345e931fa`). It took 272 ms at the
  Office.js command boundary.
- The first post-calculation sample was 904,352 KiB, about 494.6 MiB above
  baseline. Roughly 149 seconds later RSS had fallen to 504,960 KiB but was
  still 107,104 KiB (about 104.6 MiB) above baseline.
- The 9x9 board links and existing Naked-single hint/trace outputs were identical
  before and after calculation. No second calculation was run.
- Result: the existing Easy engine, which has no Master-only candidate-chain
  anchor, independently produces substantial temporary and retained growth.
  The issue is therefore not specific to the new prototype chain.
- Next, measure `_Engine Medium`, `_Engine Hard`, and `_Engine Expert` one per
  clean process using their `A1:T100` used ranges. Only after that sheet-level
  map should a formula-range subdivision or implementation change be proposed.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in `diagnostics/easy-engine-used-range-memory-20260921/SUMMARY.md`.

## 2026-09-21 Medium engine released its temporary allocation

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy was not saved.
- The final three settling RSS samples were 520,032, 520,144, and 520,144 KiB;
  their median, 520,144 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on
  `_Engine Medium!A1:T100` (command
  `cmd_e_557b60ea6c788191995a225c4719e160`). It took 116 ms at the Office.js
  command boundary.
- The first post-calculation sample was 609,072 KiB, about 86.8 MiB above
  baseline; the highest observed sample was 619,712 KiB, about 97.2 MiB above.
  Roughly 175 seconds after the first sample, RSS was 478,000 KiB, about
  41.2 MiB below baseline.
- The 9x9 board links and existing Naked-single hint/trace outputs were identical
  before and after calculation. No second calculation was run.
- Result: unlike Easy, Medium fully released its substantially smaller temporary
  allocation within the observation window. Common sheet dimensions and shared
  engine architecture are not sufficient to explain retention; the evaluated
  puzzle/formula path matters.
- Next, measure `_Engine Hard` and `_Engine Expert` one per clean process. Then
  compare the retaining and bounded sheets before subdividing formula ranges.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in
  `diagnostics/medium-engine-used-range-memory-20260921/SUMMARY.md`.

## 2026-09-21 Hard engine released its temporary allocation

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy was not saved.
- Startup RSS descended in steps before the final three baseline samples of
  533,120, 533,136, and 533,152 KiB; their median, 533,136 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on
  `_Engine Hard!A1:T100` (command
  `cmd_e_a405743512388191a20e516515b1f862`). It took 142 ms at the Office.js
  command boundary.
- The first post-calculation sample was 787,632 KiB, about 248.5 MiB above
  baseline; the highest observed sample was 794,112 KiB, about 254.9 MiB above.
  RSS crossed below baseline roughly 175 seconds after the first sample, and at
  roughly 195 seconds it was 425,648 KiB, about 105.0 MiB below baseline.
- The 9x9 board links and existing Naked-single hint/trace outputs were identical
  before and after calculation. No second calculation was run.
- Result: like Medium, Hard fully released its temporary allocation, although
  its peak was substantially larger. This further narrows retained growth away
  from shared engine dimensions and toward puzzle-specific evaluated paths.
- Next, measure `_Engine Expert` in one clean process. Then compare the four
  existing engines before subdividing the retaining sheets.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in
  `diagnostics/hard-engine-used-range-memory-20260921/SUMMARY.md`.

## 2026-09-21 Expert engine released its temporary allocation

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy was not saved.
- Startup RSS descended in a late step before the final three baseline samples
  of 400,432, 400,608, and 400,544 KiB; their median, 400,544 KiB, is the
  baseline.
- Exactly one documented `Range.calculate()` call was made on
  `_Engine Expert!A1:T100` (command
  `cmd_e_3aaa0f8eb60881918114f2f51faa611a`). It took 136 ms at the Office.js
  command boundary.
- The first and highest post-calculation sample was 689,888 KiB, about
  282.6 MiB above baseline. RSS crossed below baseline roughly 261 seconds
  later; the final 379,488 KiB sample was about 20.6 MiB below baseline.
- The 9x9 board links and existing Hidden-single hint/trace outputs were
  identical before and after calculation. No second calculation was run.
- Result: like Medium and Hard, Expert fully released its temporary allocation,
  although it had the largest peak and longest release tail of those three.
  Across the four existing `A1:T100` engines, only Easy retained material
  growth within its observation window.
- Next, compare the four engines' input states and evaluated hint/candidate paths
  before subdividing the retaining sheet or proposing an implementation change.
- After explicit action-time approval, the measurement workbook was closed with
  `Don't Save`, Excel was quit, and its main process was confirmed absent. The
  measurement copy remained byte-identical to the preserved prototype
  (`cmp=0`); production and the preserved prototype retained their hashes. Full
  evidence is in
  `diagnostics/expert-engine-used-range-memory-20260921/SUMMARY.md`.

## 2026-09-21 Existing-engine comparison narrowed the next split

- The four existing `A1:T100` engine sheets each contain 274 formula cells,
  including 192 array-formula cells. After normalizing their page and engine
  references, no formula-coordinate difference remained.
- Easy retained about 104.6 MiB at the end of its observation window. Medium,
  Hard, and Expert all crossed below baseline despite peaks ranging from about
  97.2 to 282.6 MiB. Peak allocation size therefore does not predict retention.
- Clue count, initial candidate count, and displayed hint type also fail to
  separate Easy. Easy, Medium, and Hard all display Naked singles, while only
  Easy retained.
- An independent algorithmic comparison found 17 potential elimination steps
  for Easy versus 8-10 for the other three. This is the strongest static
  difference, but it is only a candidate-topology proxy: the existing engine
  formulas do not invoke the iterative `SDKP_Chain`.
- The common hint/probe area repeatedly evaluates candidate scans through
  `SDK_CoachHintV2`, `SDK_PlayerHint`, and two independent
  `SDK_PlayerTrace` calls. The next clean-process split is therefore
  `_Engine!A61:T100`. If it retains, split `A61:T85` from `A86:T100`; otherwise
  test `A1:T59`.
- Full evidence is in
  `diagnostics/existing-engines-comparison-20260921/SUMMARY.md`.

## 2026-09-21 Easy hint/probe band released its temporary allocation

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy has not been saved.
- The final three settling RSS samples were 374,576, 374,528, and 374,608 KiB;
  their median, 374,576 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on
  `_Engine!A61:T100` (command
  `cmd_e_527775590824819187a24a56a268997f`). It took 169 ms at the Office.js
  command boundary.
- The first and highest post-calculation sample was 524,688 KiB, about
  146.6 MiB above baseline. RSS crossed below baseline roughly 99 seconds
  later; the final 365,600 KiB sample, at roughly 153 seconds, was about
  8.8 MiB below baseline.
- The hint, trace, dead-end, and diagnostic-probe outputs in `B62:B66`,
  `B69:B85`, and `B88:B100` were identical before and after calculation. No
  second calculation was run.
- Result: the combined hint/probe band fully releases its temporary allocation
  and does not explain the roughly 104.6 MiB retained after calculating the
  full Easy engine range.
- The next fresh-process test is `_Engine!A1:T59`. If it retains, split the
  live-board links (`A1:T13`), candidate display (`A14:T23`), and remaining
  constants/conflict map (`A24:T59`) one calculation per clean process.
- After explicit action-time approval, the measurement workbook was closed
  with `Don't Save`, Excel was quit, and its main process was confirmed absent.
  The copy remained byte-identical to the preserved prototype; production and
  the preserved prototype retained their hashes. Full evidence is in
  `diagnostics/easy-hintband-memory-20260921/SUMMARY.md`.

## 2026-09-21 Easy base band also released, but only after a long delay

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy has not been saved.
- After a late startup release, the final three baseline RSS samples were
  365,328, 365,456, and 365,904 KiB; their median, 365,456 KiB, is the
  baseline.
- Exactly one documented `Range.calculate()` call was made on
  `_Engine!A1:T59` (command
  `cmd_e_a3e2c9c46274819189931dd85c53472a`). It took 1,467 ms at the Office.js
  command boundary.
- The first post-calculation sample was about 423.3 MiB above baseline and the
  observed peak was about 425.1 MiB above. At roughly 99 seconds, RSS was still
  about 103.9 MiB above baseline.
- A second delayed release crossed below baseline only at roughly 161 seconds.
  The final 345,264 KiB sample at roughly 215 seconds was about 19.7 MiB below
  baseline.
- The live board, candidate display, and conflict map were identical before and
  after calculation. No second calculation was run.
- Result: `A1:T59` and the separately tested `A61:T100` both fully release when
  calculated alone. The earlier full Easy run ended at about 149 seconds with
  104.6 MiB apparently retained, only 12 seconds before the base-band test's
  second release. Its observation window may therefore have been too short.
- Do not subdivide `A1:T59` yet. Repeat `_Engine!A1:T100` once in a fresh clean
  process and observe for at least six minutes after its first post-calculation
  sample. This will distinguish delayed release from a two-band interaction.
- After explicit action-time approval, the measurement workbook was closed
  with `Don't Save`, Excel was quit, and its main process was confirmed absent.
  The copy remained byte-identical to the preserved prototype; production and
  the preserved prototype retained their hashes. Full evidence is in
  `diagnostics/easy-baseband-memory-20260921/SUMMARY.md`.

## 2026-09-21 Long repeat withdrew the Easy range-retention claim

- A fresh byte-identical prototype copy was opened in a clean Excel process.
  No cell or workbook property was edited and the copy has not been saved.
- Startup memory continued releasing for more than seven minutes. The final
  three baseline samples were 362,592, 362,416, and 362,784 KiB; their median,
  362,592 KiB, is the baseline.
- Exactly one documented `Range.calculate()` call was made on the full
  `_Engine!A1:T100` range (command
  `cmd_e_7444d5ef505881918741dc6b40deed3f`). It took 1,563 ms at the Office.js
  command boundary.
- The first and highest available sample was about 391.9 MiB above baseline.
  RSS crossed below baseline after roughly 62 seconds, remained well below it
  for the rest of the observation, and finished about 79.3 MiB below baseline
  after roughly 406 seconds.
- The board links and Naked-single hint/trace outputs were identical before and
  after calculation. No second calculation was run.
- Result: the earlier Easy endpoint of about 104.6 MiB above baseline at about
  149 seconds is not reproducible and cannot support a durable-retention or
  formula-bug conclusion. The full range and both independently tested bands
  all release fully; transient peak size and release timing vary substantially.
- The remaining relevant control is the workbook-wide application Full
  calculation, whose earlier observation stopped after about 91 seconds while
  still roughly 326 MiB above baseline. Microsoft documents this operation as
  marking all cells dirty and recalculating every open workbook. Repeat it once
  in a clean process with only the isolated copy open and observe at least six
  minutes before proposing formula changes.
- After explicit action-time approval, the measurement workbook was closed
  with `Don't Save`, Excel was quit, and its main process was confirmed absent.
  The copy remained byte-identical to the preserved prototype; production and
  the preserved prototype retained their hashes. Full evidence is in
  `diagnostics/easy-full-long-memory-20260921/SUMMARY.md`.

## 2026-09-21 Application Full long retry hit a connection timeout

- A fresh byte-identical prototype copy was opened in a clean Excel process as
  `sudoku-application-full-long-memory.xlsx`. No workbook edit or calculation
  was performed.
- After one pane reload, the intended workbook registered as connected session
  `bp_arc_e_6ab0daa450088191b37f55bf4c336049`.
- The minimum metadata command
  `cmd_e_f74eaad703a88191b6da674a741bdbd0` timed out twice, including the one
  permitted exact retry with the same idempotency key. Excel was idle at 0.0%
  CPU, so this was a connected-session timeout rather than a calculation stall.
- No Office.js Full calculation was attempted and no memory result is valid for
  this round. The copy remained byte-identical to the preserved prototype;
  production and the prototype retain their recorded hashes.
- After explicit action-time approval, the copy was closed with `Don't Save`,
  Excel's main process was confirmed absent, and the copy's hash and bytewise
  identity with the preserved prototype were reverified.
- Resume with a fresh byte-identical copy and the ChatGPT pane. Require one
  minimum metadata read to succeed, then run the planned single Application
  Full calculation with at least six minutes of observation.
- Full evidence is in
  `diagnostics/application-full-long-memory-20260921/SUMMARY.md`.

## 2026-09-21 Application Full fresh-process retry reproduced the timeout

- The prior timed-out copy was closed with `Don't Save`, Excel was confirmed
  absent, and a new byte-identical prototype copy was opened in a new Excel
  process with only that workbook present.
- The signed-in ChatGPT pane registered the intended workbook as new session
  `bp_arc_e_6ab0dc8c5f8c819190932374ddfbecb6`.
- The first minimum metadata read timed out as command
  `cmd_e_b91743151d7c8191b13ea479364b6397`. No workbook read, write, or
  calculation preceded it.
- Excel still showed `Ready`; PID 53790 used 0.0% CPU and 471,408 KiB RSS. The
  pane again lost its chat contents after the timeout. The open copy remained
  byte-identical to the preserved prototype, and production remained unchanged.
- Do not retry the same cold-start command loop without an external change to
  the Excel add-in/live-control runtime. The Application Full measurement is
  still unrun; this result is a connection/runtime reproduction, not a workbook
  memory result.
- After explicit action-time approval, the retry copy was closed with `Don't
  Save`, Excel's main process was confirmed absent, and the copy still matched
  the preserved prototype byte-for-byte. Production retained its recorded
  SHA-256.

## Why the prototype is needed

Current strategies regenerate basic candidates from the board on every call.
Pointing/Claiming/Naked Pair can describe an elimination, but subsequent
strategies and the candidate display never consume that reduced state.
Source markers were fixed by adding source coordinates before `Remove` in
four shared named formulas. Trace scalarization was fixed separately.

The assessed design rebuilds a deterministic candidate chain from the current
board. Each step consumes the preceding candidate matrix, records its
technique/sources/targets, and removes at least one actual candidate.
Stop when a valid next assignment exists, no supported elimination remains,
a contradiction occurs, or an explicit iteration budget is reached. A budget
exit must not be described as strategy exhaustion. Do not use circular
references, saved exclusion history, or automatic placement into user cells.

Deletion must recompute from the changed board and restore exclusions whose
reason no longer holds. Solver/Verify can continue to use basic candidates.
Keep the advanced candidate result shared by display/Hint/Trace rather than
rerunning the chain for each consumer. Existing engines compute B89 and B90
independently and do not gate them behind Off; avoid amplifying that cost.

## Existing offline evidence

`probe.py` in this directory is a read-only Python algorithm probe. It reads
the production workbook but never writes it. Run with the bundled Python:

```
/Users/jiangyuan/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 Sudoku/dev-notes/candidate-chain-prototype/probe.py
```

It checked 500 solution-consistent boards with full elimination to stability,
all seven advanced technique variants, 3,352 candidate removals, and a maximum
24 effective elimination steps. No certified-solution digit was removed.
Those random boards all had immediate singles; they do not independently
establish teaching-chain coverage. Ten additional Master walks using singles,
advanced eliminations, and explicitly labeled certified-solution reveals
exercised 38 advanced states with up to five elimination steps. Four eliminated
candidates in the main fixture were separately forced into the board and
rejected by an independent exact solver.

Python evidence establishes algorithm behavior, NOT Excel formula compatibility,
native recalc time, or memory stability. Reveals remain necessary for some boards.

## Main reachable Master fixture

Start with the shipped Master givens and temporarily set:

| Logical cell | Excel center | Value |
| --- | --- | --- |
| R8C3 | K26 | 1 |
| R1C2 | G5 | 6 |
| R2C3 | K8 | 4 |
| R5C3 | K17 | 3 |
| R7C3 | K23 | 6 |
| R2C1 | C8 | 5 |

Mode is `05 Master!AS11`. The desired teaching chain is:

1. Pointing row: R3C1 and R3C2 confine 7; remove 7 from R3C8 (2347 -> 234).
2. Naked pair row: R2C4 and R2C6; remove 1 from R2C7.
3. Hidden single in column: R6C7 = 1. Stop for user entry.

If computing full closure instead, a third elimination (Claiming row) removes
1 from R3C5 and R3C6. This changes the first selected basic hint, so distinguish
the two policies in tests.

## Native execution constraints and next steps

1. Treat the saved workbook as a prototype artifact only. Do not copy its
   `SDKP_` names or `_Engine Master` formulas into `Sudoku/sudoku.xlsx` yet.
2. The matched phase test is complete: normal input recalculation stayed small,
   while explicit full calculation caused repeatable large RSS peaks and retained
   growth. Do not run a long stress loop until this is isolated.
3. The unmodified prototype, the static-value anchor, and the blank anchor all
   retained roughly 289--326 MiB after one full calculation. Editing or saving
   `L102` is therefore not required. Earlier component variants were confounded
   by the same whole-application calculation and must not be treated as clean
   component isolation.
4. `_Engine Master!A1:T111` alone reproduced about 85 MiB of retained growth
   after one range calculation, versus about 326 MiB after application-level
   full calculation. Do not run additional anchor-content variants.
5. The full Easy `A1:T100` long repeat crossed below baseline after about 62
   seconds and stayed below through more than six minutes, so withdraw the
   earlier range-retention claim. All four existing engines and both Easy bands
   release their range-calculation allocations. Do not subdivide or rewrite
   their formulas based on transient RSS. Next repeat exactly one application-
   level Full calculation with only the isolated workbook open and observe for
   at least six minutes; its earlier 91-second endpoint remains the unresolved
   performance control.
6. If the memory behavior is brought within an acceptable bound, extend
   contradiction tests to malformed candidate matrices and perform a longer
   bounded delete/replay test.
7. Decide whether to keep recursive named LAMBDAs or generate an equivalent
   fixed-depth chain before production.
8. Only after that review, plan the production rollout across all five engine
   sheets, update regression tests, and obtain explicit approval before editing
   or committing the production workbook.

## 2026-09-21 structured Pointing-row vertical slice

- The user reprioritized the project toward functional completeness,
  correctness, and reasonable formula boundaries. Performance work is deferred
  unless normal use becomes unresponsive or directly exceeds resources. Do not
  resume the application-level Full-calculation investigation as the default
  next task.
- A new durable working copy was created at
  `sudoku-candidate-chain-structured-slice.xlsx`. Production
  `Sudoku/sudoku.xlsx` was not opened or edited and retained SHA-256
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.
- TDD RED was observed for two persistent `_Tests` contracts before the new
  names existed:
  1. Pointing-row detection must return separate technique, source, removal
     digits, targets, and display-text fields.
  2. Candidate application must use the removal and target fields even when the
     display text contains none of the legacy `Remove` / `from` delimiters.
- `SDKP_PointingRowStep(candidates)` now returns a five-field row:
  `technique | sources | digits | targets | displayText`.
  `SDKP_ApplyStep(candidates, step)` reads only fields 3 and 4. The old
  `SDKP_PointingRow` is now a compatibility wrapper that returns field 5.
- `SDKP_ChainLoop` uses the structured step and `SDKP_ApplyStep` when a
  Pointing-row elimination exists. The other six advanced variants still use
  the legacy text-returning path; this is intentionally a single vertical slice,
  not a partial claim that the whole chain is migrated.
- The permanent Master fixture regression still produces `R3C8 = 234`,
  `R2C7 = 67`, status `ASSIGNMENT`, step count `2`, and next hint
  `Hidden single in column · R6C7 = 1`.
- A native delete/replay check passed after using an actual clear operation:
  clearing `05 Master!K26` restored R3C8 to `2347`, R2C7 to `167`, and the
  opening hidden single; restoring K26 replayed the two eliminations.
- The saved working copy was returned to a clean start: all six fixture inputs
  are blank, `05 Master!AS11` is `Off`, and the Master live-board count equals
  its 23 givens. The stale engine-wiring regression was updated to accept the
  four standard `SDK_PlayerHint` paths plus the Master shared-chain fallback.
- Saved native regression state: **96 PASS, 0 FAIL, 18 SKIP**. The three new
  structured-slice tests are rows 165--167 and all pass. Offline package
  inspection confirmed 19 `SDKP_` names, the two new names, the structured
  chain dependencies, and SHA-256
  `65bc2b1dffb33936971524a53b5ad7f4afb6ae4f305956b1a8ed04a71d093c47`.
- Excel's `AND` unexpectedly returned `FALSE` when directly aggregating five
  individually true `INDEX` comparisons against the dynamic five-field row.
  The regression uses an explicit `SUM(--HSTACK(...))=5` assertion; the five
  individual comparisons were inspected and were all true.

### Current next step

1. Keep production unchanged.
2. The Pointing Row/Column pair is now structurally migrated; see the
   2026-09-22 checkpoint below.
3. Add duplicate-forced-single and no-progress/error guards before migrating
   claiming or naked-pair strategies. Write native failing tests first.
4. Do not expand to all five engine sheets until every migrated variant has a
   permanent native regression and the working copy saves with zero failures.

## 2026-09-22 structured Pointing-column checkpoint

- Continued in the independent
  `sudoku-candidate-chain-structured-slice.xlsx` working copy. Production
  `Sudoku/sudoku.xlsx` was not opened or edited and retained SHA-256
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.
- Added a persistent `_Tests!168` five-field Pointing-column contract test
  before changing the names. It was RED (expected TRUE, actual FALSE, result
  FAIL) because `SDKP_PointingColStep` did not yet exist.
- Added `SDKP_PointingColStep(candidates)` with the same
  `technique | sources | digits | targets | displayText` contract and
  deterministic first-hit selection as Pointing Row. The old
  `SDKP_PointingCol` now returns only field 5 as a compatibility wrapper.
- `SDKP_ChainLoop` selects structured Pointing Row first, structured Pointing
  Column second, and then the legacy advanced dispatcher. Either structured
  step is applied by `SDKP_ApplyStep`, using its digit and target fields rather
  than parsing display prose. Claiming and Naked Pair remain on the legacy
  text path; no production rollout has occurred.
- The GREEN native regression was **97 PASS, 0 FAIL, 18 SKIP**. Rows 165--168,
  the existing public Pointing-column wording test (row 118), and the Master
  fallback test (row 142) passed. The permanent Master chain fixture (row
  167) still returned R3C8 `234`, R2C7 `67`, status `ASSIGNMENT`, two steps,
  then `Hidden single in column · R6C7 = 1`.
- The workbook was saved in place through Excel's official save API after a
  direct live count of 23 digits in `_Engine Master!B3:J11`. The six Master
  fixture inputs were blank and `05 Master!AS11` was `Off`. The saved XLSX
  passed ZIP integrity and offline cached-value checks: 97 PASS, 0 FAIL,
  18 SKIP, 20 `SDKP_` names, 23 live digits, and both new structured names.
  Saved SHA-256:
  `8a409d18e5598892054a48c69d7ae22dfe4380800f9b777f491c7f9907b015aa`.
- A broad Office.js read-only follow-up hit `RangeAreas` readback limitations
  and was stopped. Small direct range reads then confirmed the test rows,
  clean inputs, mode, and no formula-error values in the affected test ranges.
  The ChatGPT add-in reached its own usage limit before the save, but the
  connected Excel session remained usable for direct count and save calls.

### Next bounded round

1. Keep this saved Pointing pair as the durable baseline. Do not resume
   application-level Full-calculation memory experiments by default.
2. Write a failing native test for duplicate-forced-single contradiction
   detection, then implement and verify the smallest guard in the prototype.
3. Separately test a no-progress elimination (same candidate count before and
   after apply) and return an explicit error/stop state rather than recurse.
4. Only then consider another advanced-strategy migration or production
   rollout. Production remains untouched until explicit approval.

## 2026-09-22 duplicate forced-single contradiction checkpoint

- Continued only in the independent
  `sudoku-candidate-chain-structured-slice.xlsx` prototype. Production
  `Sudoku/sudoku.xlsx` was not edited and retained SHA-256
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.
- Added persistent native `_Tests!169:172` cases before changing formulas:
  duplicate forced digit 1 in a row, a column, and a box (but not a shared
  row/column), plus a separated-singleton control. The first three were RED
  (`FAIL`), while the control passed.
- Added `SDKP_ForcedConflict(board,c)` and included it in
  `SDKP_Contradiction`. It checks only blank cells with exactly one candidate,
  then detects duplicate forced digits in any row, column, or box. The
  strategy order, display wording, live board, and production workbook were
  not changed.
- The GREEN Excel-native regression returned **101 PASS, 0 FAIL, 18 SKIP**.
  All four new cases and the pre-existing structured-chain fixture passed.
  The name definitions were read back from Excel. The saved XLSX passed ZIP
  integrity and offline cached-result checks: 101 PASS, 0 FAIL, 18 SKIP,
  21 `SDKP_` names, all four new cases PASS, and the new guard referenced by
  `SDKP_Contradiction`.
- The clean Master board still has 23 givens; `05 Master!AS11` is `Off`, and
  all six temporary fixture input cells are blank. The prototype was saved
  in place through Excel. Saved SHA-256:
  `8d6194fc10d327c194fc8d933a21901584e39cd8655d19a6cf808588b7596ca3`.

### Next bounded round

1. Use this saved prototype as the baseline; do not alter production yet.
2. Test a no-progress elimination where applying a step does not reduce the
   candidate count. Add an explicit stop/error state instead of recurring.
3. Recheck the native suite and clean Master state before considering another
   strategy migration or rollout. Keep performance experiments deferred unless
   normal use is blocked or resources are exceeded.

## 2026-09-22 no-progress elimination checkpoint

- Continued in the saved independent structured-slice prototype from commit
  `21bd9ad`. Production `Sudoku/sudoku.xlsx` was not edited and retained
  SHA-256
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`.
- Root cause: `SDKP_ChainLoop` previously recurred after any nonempty advanced
  hint without checking whether applying that hint actually shortened the
  candidate matrix. A no-op could consume the step budget and end as `LIMIT`
  rather than identify the faulty elimination.
- Added persistent native `_Tests!173:175` before implementation: removing
  an absent digit is not progress, removing a present digit is progress, and
  growing a candidate set is not progress. All three were RED (`FAIL`) before
  `SDKP_HasProgress` existed and GREEN (`PASS`) afterward.
- Added `SDKP_HasProgress(before,after)` to compare total candidate-string
  lengths. `SDKP_ChainLoop` now recurs only when this total decreases. On a
  nondecreasing result, it returns the original candidate matrix with status
  `NO_PROGRESS` and the attempted hint, without increasing the step count.
  This is a defensive guard; the currently valid strategy functions normally
  generate an effective elimination.
- The Excel-native regression was **104 PASS, 0 FAIL, 18 SKIP**, including the
  Master two-step chain fixture and all three new tests. Excel readback
  confirmed the guard is in the recursive branch. The current strategy set
  does not naturally generate a no-op, so the terminal `NO_PROGRESS` branch
  has not been triggered end-to-end in a native chain test; the predicate and
  existing normal recursion path were tested separately.
- The prototype was saved in place through Excel. Its ZIP passed integrity
  checking and offline cached-result verification: 104 PASS, 0 FAIL, 18 SKIP,
  22 `SDKP_` names, and all three new test results PASS. The Master live board
  still has 23 givens, the six temporary fixture inputs are blank, and mode
  is `Off`. Saved SHA-256:
  `7cab569e47f1976ab4d750c4add4c88de68058aff7304ff561675289044f4a8f`.

### Next bounded round

1. Keep production untouched. Decide whether a small, isolated test seam can
   exercise the `NO_PROGRESS` terminal output in Excel without altering
   strategy behavior or adding test-only parameters to the public chain API.
2. Then audit the remaining legacy text-parser path, beginning with Claiming
   Row, before migrating another strategy. Treat performance work as deferred
   unless normal use is blocked or resources exceed practical limits.

## 2026-09-22 no-progress test-seam and Claiming Row audit

- Read-only audit of the connected, saved structured-slice prototype at
  commit `43a5592`. Neither workbook was edited or saved. Production and
  prototype SHA-256 remained, respectively,
  `f69c6b84f19eb46334167b1729cee1e844e225eaf3a150b00a58eb1d7272125a`
  and `7cab569e47f1976ab4d750c4add4c88de68058aff7304ff561675289044f4a8f`.
  Native `_Tests!E1:E175` still showed **104 PASS, 0 FAIL, 18 SKIP**.
- `SDKP_ChainLoop` has no input for an externally supplied advanced step.
  The current Pointing, Claiming, and Naked Pair detectors construct targets
  from candidate cells that actually contain the digit to remove. Therefore
  a valid strategy-generated step does not naturally enter the
  `NO_PROGRESS` branch. A full-chain synthetic trigger would require a
  temporary override of a workbook strategy name, an unverified name-scope
  trick, or a new test-only injection API. None is justified merely to test
  this defensive branch; keep the existing native predicate cases and the
  verified branch readback, and disclose this test boundary before rollout.
- Direct formula audit shows `SDKP_AdvancedHint` selects `SDKP_ClaimingRow`
  before `SDKP_ClaimingCol` and Naked Pair variants when structured Pointing
  finds nothing. `SDKP_ClaimingRow` emits only display prose. `SDKP_Apply`
  parses `Remove ... from ...` back from that prose to identify digits and
  targets, so changes to display wording can change the elimination action.
- `_Tests!19` covers the legacy `SDK_HintClaiming` public wording, not the
  prototype `SDKP_ClaimingRow`; searching `_Tests!A1:E175` found no direct
  `SDKP_Claiming*` formula. This is a prototype coverage gap, not evidence of
  an observed wrong result. Pointing Row and Column already use five-field
  structured steps and apply their data fields without parsing display text.

### Next bounded round

1. Write a RED native five-field contract test for
   `SDKP_ClaimingRowStep(candidates)` using a small synthetic matrix. Place
   digit 3 only at R2C7 and R2C8 in row 2, with R3C7 as the sole box target.
2. If the RED test fails for the expected missing-name reason, implement the
   smallest Claiming Row structured step, preserve the existing public text
   wrapper, and route only that structured step through `SDKP_ApplyStep`.
   Keep Claiming Column and Naked Pair on the existing path for now.
3. Run the full native regression and clean Master checks, then save and
   checkpoint only the independent prototype. Do not alter production until
   separately approved.

## Tool pitfalls from earlier successful rounds

- Offline formula snapshots contain `_xlws.FILTER`. Office.js input requires
  plain `FILTER`; writing the storage prefix caused `#NAME?`.
- `write_range` with `value:null` reported success but did not clear inputs.
  Use advertised clear commands or `range.clear(Excel.ClearApplyTo.contents)`
  and verify the result.
- CUA `app.click` takes an integer index, not `{id:...}`. Do not invent keyboard
  method names; inspect returned API documentation.
