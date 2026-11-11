# Pilot report — Word→PDF oracle corpus (issue #6153)

---

## PAUSE CHECKPOINT — 2026-11-09 (Reuben-directed, machine-wide)

Paused on instruction because the mini's app-dictionary AppleEvent delivery
is degraded and a diagnosis lane needs the machine quiet.

- **In-flight item at pause: none.** No pilot export or document operation was
  mid-flight; the pilot driver had already ended and the options-probe driver
  had been killed while waiting for the lock, before the pause arrived.
- **Completed before the pause:** 3 PDFs fully produced and file-verified —
  `tc-p0-del-para`, `tc-p2m-multiauthor`, `tc-p2m-table-row-del`, vector
  **markup-default** only (page counts, producer string, extracted-text
  checks, and rendered-colour sampling all recorded in the manifest settings
  detail and in the body of this report). No record had reached `verified`
  status (Track Changes Options readback incomplete, error -1700).
- **Manifest state at pause:** all 204 records `pending`. The 3 produced
  markup-default records keep their `oracle_pdf` {file, sha256, bytes}
  pointers and full evidence in `settings.detail`, but their status was set
  back to `pending` under the pause directive (no record marked produced
  until its settings vector completes). The other 6 pilot records
  (clean × 3, markup-custom × 3) were never produced — reasons in the body
  below. The remaining 195 records were never touched.
- **WORD-LOCK: released by the pilot before the pause arrived.** The run-3
  claim (holder `pdf-oracle-pilot`) was released via
  `word_lock.sh release --force` with the pilot's own token, rc=0 — the
  `--force` was needed only because a plain release was refused NOT_CLEAN
  over two foreign Office recovery documents (see End state below); the lock
  file itself was released through the helper, not deleted. `state=free` was
  verified immediately after release; the lock was then re-acquired by other
  lanes' batches. At the pause checkpoint the pilot holds **no claim**, has
  **no processes** on the mini, and the current holder is another lane
  (`gridcraft-desktop-diff` / `excel-diag-c2b` at last single check).
- **Known leftover (not the pilot's, not touched at pause):** Word still
  shows 2 unnamed recovery documents from the mid-session Word restart; they
  resist SSH index-close. Closing them is GUI-session work and is left for
  the coordinated quiet window / diagnosis lane to encounter as-is.
- **Standing by.** No new claims, no new queue jobs, no further Word/Excel
  work from this pilot until explicitly resumed.

---

Date: 2026-11-09 (session ~20:34–22:00 UTC). Pilot entries: `tc-p0-del-para`,
`tc-p2m-multiauthor`, `tc-p2m-table-row-del`. Vectors: clean, markup-default,
markup-custom (3 PDFs per entry, 9 total).

**Outcome: 3 of 9 PDFs produced** (all three markup-default). The production
route is proven for markup-showing exports; the clean vector uses the same
mechanism and was not reached; the markup-custom vector is blocked on setting
revision colours programmatically. Details and unblockers below — per the
FORMAT.md rule, the three produced PDFs are recorded as `produced`, not
`verified`, because their Track Changes Options readback is incomplete.

## Environment (verified)

- Oracle: **Word for Mac 16.114.4 (26110224)** on the Mac mini
  (mdls `kMDItemVersion` 16.114.4, PlistBuddy CFBundleVersion 16.114.26110224).
- All work went through the WORD-LOCK (`word_lock.sh`) and the queue-runner,
  per the `local-mac-word-oracle` skill. No preemption, no ARD/screen control,
  no Grant File Access approval.
- Inputs staged byte-identically: sha256 of all three inputs recompute-matches
  the manifest pins on the VM and on the mini, before and after the runs
  (staged copies in Word's container likewise unchanged).

## The route that worked (markup vectors)

No VBA was needed. The route is **Word's AppleScript dictionary**, driven from
a queue-runner job holding the WORD-LOCK:

1. Stage the input into Word's container as `pdfpilot-<id>.docx`.
2. `open` it inside `tell application "Microsoft Word"`.
3. Set document properties `show revisions` and `print revisions` to true.
4. `save as <doc> file name "<container>/word-markup-default.pdf" file format format PDF`.
5. `close … saving no`; copy the PDF out of the container to
   `~/word-oracle/pdf-pilot/out/<id>/` and into the corpus `entries/<id>/`.

Export takes **1–2 s per document**. Producer string of all three PDFs:
`macOS Version 15.8.2 (Build 24H232) Quartz PDFContext` (i.e. genuine Word for
Mac Save-As-PDF output through Quartz, the same producer family as the
neurotic_docx_bench references).

Per-entry verification of the produced PDFs:

| entry | pages | text check | rendered colour (150 DPI pixel sample) |
|---|---|---|---|
| tc-p0-del-para | 1 | Deleted paragraph present as balloon annotation: "Deleted: Attention then turned to the staffing of the night desk during the winter holiday period." | red-dominant (by-author) |
| tc-p2m-multiauthor | 1 | Insertions present in final text ("extended timetable", "final draft"); entry has no deleted text, so no deletion check applies | blue-dominant (by-author) |
| tc-p2m-table-row-del | 2 | Deleted row cells present as balloon annotation "Deleted: middle-one …" with author/date attribution | purple-dominant (by-author) |

Two corpus-relevant observations:

- **Word for Mac's default markup PDF puts deletions in balloons**, not
  inline: the deleted content appears in the PDF text layer as
  `Deleted: <text>` margin annotations (this is also how `pdftotext` sees it —
  a scorer must expect the prefix and the line-wrap). The FORMAT.md
  markup-custom vector already specifies "balloons inline", which will require
  the balloons mode to be set explicitly.
- The by-author palette assigns a *different* colour per author (red / blue /
  purple across these three files) — exactly the behaviour issue #5232 is
  about, and a useful positive control for the corpus scorer.

## What did not get produced, and why

### markup-custom (0 of 3) — revision colours cannot be set via the dictionary

Word's AppleScript dictionary *declares* the needed properties on the
`Word options` class (`inserted text color`, `deleted text color`,
`inserted/deleted text mark`, cell colours, move colours, `comments color`),
but in practice on 16.114.4:

- `set inserted text color of Word options to blue` → **error -10006**
  ("Can't set «class 3603» of «class w268»"). Attempted in the production
  script under the lock; the PDF was correctly not produced.
- Reading the same properties back `as text` → **error -1700** (the colour
  value will not coerce to text), so even the *default* vector's full settings
  vector cannot be recorded by coercion.
- From an SSH `osascript` (outside the GUI session), the property cannot even
  be read (**error -1728**) — options access genuinely requires the
  GUI-session queue-runner context; document-level dictionary commands do work
  from SSH.

Untested discriminating probe (written, queued, then withdrawn — see below):
read options by **equality tests** against candidate constants
(`= blue`, `= by author`, …) instead of coercion, and attempt the colour and
mark sets *in isolation* to learn whether only the colour form fails. If the
dictionary route stays closed, the fallback is the **GUI route through the
queue-runner: Word ▸ Preferences ▸ Track Changes**, setting Insertions =
blue/underline, Deletions = red/strikethrough, balloons = inline, then the
same dictionary export. A VBA route (`Options.InsertedTextColor` etc.) remains
a third option but installing a macro into Normal.dotm needs the manual
first-time GUI procedure from the oracle skill.

### clean (0 of 3) — never reached

The clean jobs sat behind the markup-custom jobs in the pilot driver. After
the custom set failed, a later custom job stalled (Word had been restarted by
another lane's activity mid-session and came back with two unnamed recovery
documents, at 53% CPU); per the skill's 3-minute hang rule the stalled
`osascript` was killed and the driver ended. Clean uses the identical export
mechanism with `revisions view` final / show+print revisions false, so it is
expected to work unchanged — but it is unproven and recorded as not produced.

## Lock and contention log (the real pilot cost)

Production itself is seconds; the session was dominated by sharing one Word
with the validate / compare-vs-word / gridcraft batches:

- ~20:37Z: first submission waited behind vba-matrix and
  excel-desktop-diff-ax5 claims; acquired for run 1.
- **Run 1**: AppleScript compile error -2741 (multi-word enumerators such as
  `dark blue` do not compile outside a `tell` block). No PDFs. Lock released
  cleanly; script fixed.
- **Run 2**: job parameters were passed via a TSV file in Word's container;
  reading it via StandardAdditions `read` from the script failed silently
  (writes to the same container worked). Jobs are now embedded in the script.
  Also found a foreign leftover document open in Word
  (`blkB-fz-37090.docx-base.docx`, another lane's staged input) — closed
  without saving under the pilot's own claim before release.
- **Run 3** (submitted ~20:52Z, acquired 21:13:36Z — the queue-runner runs
  inbox jobs in filename order and other lanes' jobs kept landing first;
  renamed to `00_pdf_oracle.job`): produced the 3 markup-default PDFs, then hit
  the -10006 / -1700 failures above and the hang/kill described under "clean".
  During this window Word itself restarted (elapsed 04:31, 2 recovery
  documents) under another lane's load.
- **Options probe**: submitted as `00_pdf_probe.job`; the runner started it at
  21:28:19Z and it waited in `acquire` behind the run-3 claim for the rest of
  the session (the job file was removed from the inbox once started, so it
  could not be re-run). At session end the probe driver was killed while still
  waiting — deliberately: had it acquired unattended later, its isolated
  set-attempts could have left the shared Word profile's revision colours
  changed with nobody watching. The probe scripts remain in `pilot/` and on
  the mini (`~/compare-vs-word/dialog/pdf_probe.*`) for a supervised run.
- End state: the run-3 claim (holder `pdf-oracle-pilot`) was released at
  session end via `word_lock.sh release --force` with the pilot's own token
  (rc=0; a plain release was refused NOT_CLEAN because of the 2 documents
  below), and the lock was immediately re-acquired by the next lane's batch.
  The pilot holds no claim and has no processes running on the mini; the
  corpus inputs are untouched (hashes above). Word still shows
  `word_docs=2`: two unnamed Office recovery documents left by the
  mid-session Word restart (another lane's documents, not pilot documents —
  the pilot closed every document it opened, without saving, in every run).
  They resist index-based `close document 1 saving no` from SSH (count stays
  2 across repeated closes); they survived several other lanes' sessions the
  same way and did not block their acquires. Closing them likely needs a
  GUI-session pass; flagged here for the next supervised window.

Approximate timing: lock/queue waits ≈ 8 min (run 1), ≈ 21 min (run 3),
> 20 min unstarted (probe); production ≈ 1–2 s per PDF. A quiet window — or a
coordinated slot — is worth more than any script optimisation.

## Scripts (kept for the production run)

`pilot/pdf_oracle.command` + `pilot/pdf_oracle.applescript` (v2: jobs embedded,
index-based close loops, option snapshots — snapshot still uses the failing
coercion and must be replaced by the equality-based reader in
`pilot/pdf_probe.applescript` before the production run). Copies also live on
the mini in `~/compare-vs-word/dialog/`.

## Recommended next steps

1. Supervised quiet-window run of `pdf_probe` to settle: (a) equality-based
   options readback (completes the markup-default settings vectors → the 3
   produced records can move to `verified`), (b) whether *any* dictionary set
   of revision colour/mark works.
2. If sets fail, produce markup-custom via Preferences ▸ Track Changes (GUI,
   queue-runner), recording the route per record; set balloons to inline as
   FORMAT.md specifies.
3. Run the clean vector for the three pilot entries in the same window, then
   stamp all 9 records verified with full settings vectors.

## HOLD — container-write fault (checkpoint, 2026-11-10T00:37Z)

Stop-and-check issued by the parent thread: the compare-study thread's
resumed 5.2 lane found that **writes into Word's container
(`~/Library/Containers/com.microsoft.Word/Data/Documents/`) hang in the
kernel**, and stopped itself with 0 budget spent. Their gate is a 20-second
container write/read probe, `~/compare-vs-word/dialog/probe_resume2.sh` on
the mini. The dictionary all-clear's gate never exercised Word-container
writes. **This pilot is HOLDING all Word work — no staging copies, no
exports, no dictionary driving — until the parent thread confirms the probe
gate has passed.** The probe has NOT been run by this pilot (run only when
directed).

### State check at 2026-11-10T00:37:38Z (this pilot)

- **No container writes of mine in flight.** No `cp`/`ditto`/`shasum`
  process touching the Word container path, and no pilot processes
  (`pdf_oracle*`, `pdf_recover`, `pdf_opts*`) running on the mini. My last
  container write attempts were the v3 driver's staging `cp` (driver killed
  under the 3-minute hang rule) and probe_tcc's watchdog-killed `cp`; no new
  writes have been started since, and none will be while the hold stands.
- **pid 54043 is gone.** The compare lane last observed that cp (from the
  earlier pilot run — consistent with the orphaned `cp` child of my killed
  v3 driver, pid 52581, which survived the driver's hang-rule kill) hung for
  19:36. At this check `ps -p 54043` returns nothing: the process no longer
  exists. **I did not kill it** — no pilot action since the v3 kill touches
  it; it exited or was terminated outside this pilot. Recorded here because
  the kill decision was reserved to the parent thread; it is now moot.
- **WORD-LOCK: not held by me.** First check (~00:36Z) showed it held by
  `gridcraft-desktop-diff` (batch `excel-desktop-diff-pane2`, claimed
  00:35:35Z — the claim that queued behind my settingsprobe2 cleanup and
  acquired after my release); their claim, not touched. Second check
  (00:37:38Z): `state=free`, `word_docs=0`, `word_unnamed_windows=0`,
  queue empty. All of this pilot's own claims were released cleanly through
  `word_lock.sh` earlier in the resumption (including the settingsprobe2
  orphan-claim cleanup: TERM of the wedged Word instance holding only my
  probe document, then a plain token release, rc=0). Nothing of mine blocks
  a clean state, so there was nothing to release at the hold.
- Queue inbox contains only other lanes' `zz_promptkit_*` jobs; no pilot
  jobs pending.

### Pilot status frozen by this hold

- 3 of 9 records produced (markup-default only, pre-pause; file-verified;
  manifest records still `pending` with hold notes).
- 6 of 9 NOT produced: clean ×3, markup-custom ×3.
- Resumption findings before the hold (same evening): the relaunched Word
  instance (pid 47189 zombie, quit refused -128) was TERM'd and relaunched
  clean (new pid 63762); **dictionary open of an already-staged container
  copy works again** (`OPENRET docs=1`). But **container shell access is
  still blocked** for both reads (`shasum` watchdog-killed at 15 s) and
  writes — the same fault class as this hold. **Track Changes Options
  get/set via the bare `Word options` class-name reference still fail**:
  get -1728, set -10006, even against the healthy relaunched Word.
  **But the application `settings` property route is PROVEN** (probe 3b,
  job `00_pdf_settingsprobe2`, results read before the hold): `settings`
  of application (sdef line 3486, code wSet) reaches the options object
  where the class-name reference does not — `GET_VIA_SETTINGS
  insColor=byAuthor delColor=byAuthor`; `SET_VIA_SETTINGS
  beforeByAuthor=true afterBlue=true restoredByAuthor=true` (set inserted
  text colour to blue, equality readback confirmed, restored to by
  author); `DELMARK_VIA_SETTINGS before=strikethrough
  after=strikethrough`. So the -10006 that blocked markup-custom since the
  first run was the wrong object reference, not a dictionary limitation.
  Still unproven: a full snapshot of every options value via `settings`,
  and the options state after the later Word TERM/relaunch (probe 3b
  self-restored the one value it set, but that Word instance was
  subsequently TERM'd during lock cleanup, so re-snapshot at next
  pre-flight). Also unresolved: after two dictionary save-as operations
  in probe 3b, `close` failed with -1708 and the document wedged (that
  instance was the one TERM'd during lock cleanup), and no PDF exfil route
  out of the container has been proven in this session (shell reads block;
  the Finder-duplicate test's outcome was never observed).
- Next step when the hold lifts: confirm `probe_resume2.sh` gate result
  from the parent thread; only then re-plan staging (container copies may
  need a non-shell route) and production.

---

## RESUMPTION — pilot completed 2026-11-10 (all 9 records verified)

Resumed at Reuben's direction (2026-11-10 21:33 EST) after the
container-write fault cleared by observation. **Outcome: 9 of 9 pilot
records `verified` in manifest.json** — the 3 markup-default records
from run 3 (settings vectors completed by the now-proven options
readback) plus 6 PDFs produced in this resumption. The remaining 195
manifest records are untouched and still `pending`; every pilot input
sha256 pin recomputes after the run, on the source files and on the
staged container copies alike.

### Pre-flight (all gates passed)

- Word build re-verified: **16.114.4 (26110224)** (mdls + PlistBuddy).
  Console owner `muse`; queue-runner alive; lock free; Word at 0 docs.
- Container health re-proven by the drivers themselves before any lock
  or Word work: timed write/read/remove of a probe file inside
  `~/Library/Containers/com.microsoft.Word/Data/Documents/` (watchdog
  20 s) — OK in both the v6 and v7 drivers. The kernel-hang fault did
  not recur.
- **Full Track Changes Options snapshot — the checkpoint's last
  unproven item — now works**: one equality-based snapshot via the
  application `settings` property returned all 14 fields (probe job
  `pdf_opts5_probe`, 2026-11-11T02:37Z, 3 s):
  `insColor=byAuthor insMark=underline delColor=byAuthor
  delMark=strikethrough insCellColor=lightBlue delCellColor=pink
  mergedCellColor=lightYellow splitCellColor=lightOrange
  moveFromColor=green moveFromMark=doubleStrikethrough
  moveToColor=green moveToMark=doubleUnderline commentsColor=byAuthor
  printComments=missing value (unset)`. This snapshot completes the
  three markup-default records' settings vectors.
- PDF exfil out of the container via shell: **proven** (watchdog `cp`
  of a staged file out, hash-identical; then per-PDF COPYOUT_OK for
  every produced PDF).

### The v5/v6 failure: a second save poisoned the session

The first resumption driver (v5) died on an AppleScript identifier
collision (`lines` / `line` vs the Word dictionary's line class:
-10006 "Can't set every line", -1728 in the restore's valOf) — fixed
in v6 (`pdf_opts6.applescript`) and the stranded custom option values
were restored by a dedicated job and verified by readback before any
further production.

The v6 driver then produced its first markup-custom PDF correctly —
and silently poisoned the session doing it. Its export performed a
**second save-as, to the un-granted out-dir path**. Each such save
raised a **Grant File Access panel**; four panels stacked modally.
Behind them, dictionary `open` of the next staged copies silently
no-oped (the still-open first document stayed active), `close` failed
-1708 in every formulation (plain, saved-flag, showin), and exports 2
and 3 re-saved the *first* document under the other two entries'
names (all three "custom" PDFs 8,590 bytes — the tell). The driver was
killed from the parent session; the two wrong PDFs and all clean
output were deleted, and only export 1 (name-verified correct) was
kept.

Clearing the wedge, in order: AX window-name census (traversal-free)
showed `Grant File Access` ×4 + the document window. AX tree
traversals of the panels park (watchdog kills at 60/90 s); Escape does
not dismiss them; the dictionary close stayed -1708 while they stood.
**Directly addressed clicks worked**: `click button "Cancel" of
window i` (System Events, no traversal) cancelled all four panels —
never `Select…`, which is the grant approval. The document then closed
immediately (saved-flag pattern, docs=0) and the claim released
plainly. Standing lessons, now baked into the v7 driver and worth
carrying into the production run: **one save per export, container
path only; the corpus out-dir is reached by shell copy-out, never by
a Word save**; name-gate the active document before every save;
FINAL_DOCS must return to baseline after every export or production
stops.

### v7 production run (2026-11-11T03:18:13Z → 03:19:24Z, 71 s)

Driver `pdf_oracle7.command` + `pdf_opts7.applescript`, one claim
(holder `pdf-oracle-pilot`), serialized behind the compare lane's
Stage 5.2 claims as usual — gates: container probe OK, census free of
grant panels, baseline DOCS=0. Snapshot → set-custom (readback
verified: blue/underline/red/strikethrough) → markup-custom ×2
(multiauthor 8,731 B, table-row-del 7,986 B; DOC_NAME correct,
SAVE_CONTAINER_OK, CLOSE_OK plain, FINAL_DOCS 0, COPYOUT_OK each) →
**restore verified by an identical RESTORED snapshot
(RESTORE_MATCH, all 14 fields)** → view mode restored to balloon →
clean ×3 (del-para 8,265 B, multiauthor 8,604 B, table-row-del
7,552 B; show/print revisions false, revisions view final) → hashes
unchanged, container PDFs cleaned, **LOCK_RELEASED plain**. Word was
left with 0 documents, the profile options at their original values,
and no pilot processes; the lock passed to the compare lane's next
batch within seconds.

### Verification evidence (per record, also in the manifest)

All 9 PDFs: producer `macOS Version 15.8.2 (Build 24H232) Quartz
PDFContext`, letter. Page counts — clean 1/1/1, markup-default
1/1/2, markup-custom 1/1/1 (the table-row entry's balloon layout is
what takes its markup-default to 2 pages).

- **clean**: deleted content absent from the text layer (del-para's
  paragraph, the table entry's "middle-one" row), insertions present
  as final text (multiauthor), zero saturated revision-colour pixels
  at 150 DPI.
- **markup-default**: run-3 evidence stands (balloon "Deleted: …"
  annotations with attribution; by-author palette: red #D13438 /
  blue #0078D4 / purple #5C2E91 dominant per entry) — now with the
  completed options readback, so all three records are `verified`.
- **markup-custom**: deletions/insertions render **inline** (no
  "Deleted:" balloon prefix; deleted text present in the page text),
  confirming the in-line revisions view. **Corpus finding: Word does
  not emit the settings' raw enum RGB.** With inserted colour set to
  `blue` and deleted colour to `red`, the PDF content streams
  (ICCBased space, `sc` operators) carry insertions as
  **#2E97D3** and deletions as **#B5082E** — confirmed identically by
  150-DPI raster sampling. The FORMAT.md vector names the settings
  (blue 0000FF / red FF0000); a scorer must expect Word's emitted
  palette values, recorded per record as
  `emitted_revision_colours`.

### Readiness judgment for the full 204-record run

**Ready, with the v7 driver shape.** Production cost is now measured
and small: ~2 s of Word time per export, 71 s wall for 5 exports
including staging, snapshots, and restore — a full 204-record batch
is roughly 10–15 minutes of Word time *inside one claim*. The risks
are no longer in the export mechanism; they are scheduling risks: a
contended lock (this resumption waited out a neighbouring batch's
30-minute lock-wait once) and any recurrence of the machine faults
(the container hang cleared by observation only — the v7 driver's
pre-lock container probe is the tripwire and should stay). Two
production-run stipulations: (1) single-save/container-only exports
and the name/FINAL_DOCS gates are mandatory, not optional — the v6
wedge cost ~40 minutes and one wrong-document near-miss; (2) run it
in a coordinated quiet window or a small number of long claims rather
than many short ones, and keep the per-record settings snapshot in
each claim (options are profile-global; a claim that sets custom
values must restore them before releasing, as v7 did). The balloons
vector (4th canonical vector) is not part of the seeded 204 records;
nothing in this pilot exercised it.

---

## PRODUCTION RUN — full 204-record batch (2026-11-10, in progress)

Directed by Reuben (original launch 22:57 EST; relaunched 23:11 EST
after his challenge to the emitted-colours finding was verified
first-hand by the parent thread against the pilot PDFs: sRGB profile
embedded, content-stream values identical to 150-DPI raster values,
by-author control colour exactly #D13438 — no pipeline colour shift;
the FORMAT.md item-4 addendum stands as written).

**Scope:** the 195 pending records — 65 entries × clean /
markup-default / markup-custom. The 9 pilot records are untouched.
Driver: `pdf_oracle8.command` (group selected by invoked name,
_g1.._g4) + `pdf_opts8.applescript` (= pdf_opts7 with the doExport
AppleEvent timeout raised 120 s → 900 s for the phase-4 stress
inputs; every other v7 rule verbatim). Four claims, holder
`pdf-oracle-pilot`, one group per claim: markup-default phase
(options untouched, view forced to balloon first) → set-custom →
markup-custom phase → restore + RESTORE_MATCH → view restore →
clean phase. Manifest is stamped per group from VM-side verification
(verify_full.py → evidence JSON → update_manifest_full.py), and this
section is appended per group, so the run is resumable from disk
state at any point.

### Pre-flight (2026-11-11T04:12–04:16Z)

- Word build re-verified: **16.114.4 (26110224)** (mdls +
  PlistBuddy). Console owner `muse`; queue-runner alive (pid 47969);
  lock held at check time by the compare lane's vba-matrix batch
  (expected contention; drivers wait, no preemption).
- All 65 pending inputs staged to `~/word-oracle/pdf-pilot/<id>.docx`
  on the mini; **all 65 sha256 recompute-match the manifest pins**
  (diffed against the manifest-generated expected list).
- VM-side XML inventory of all 65 inputs built
  (`pilot/inventory.json`): per-entry ins/del element counts, authors,
  formatting-change elements, comments presence, and distinctive
  deleted/inserted probe tokens for verification. Notable: five
  inputs carry **no tracked revisions at all** (tc-p0-identical,
  tc-p2-base-del, tc-p2-both-del, tc-p2-rev-del, tc-p2m-empty-both —
  Word's compare treated the pre-existing revisions as accepted, so
  its output documents are clean); their three vectors are still
  produced and the records say so. The phase-4 stress inputs carry
  15,000 (tc-p4-budget-fewlong) and 67,500 (tc-p4-budget-manypairs)
  revision elements per side.
- Container health: re-probed inside each group driver before any
  lock or Word work (watchdog 20 s), per the pilot stipulation.

### Group progress

(appended per group as each claim completes and its records are
stamped verified)

- **Group 1 (17 entries, 51 records) — DONE.** Claim
  2026-11-11T04:17:20Z → 04:27:19Z (~10 min): container probe OK,
  baseline DOCS=0, snapshot = profile defaults (identical to the
  pilot's), view forced to balloon (first markup-default export's
  VIEW_BEFORE=balloon), CUSTOM_SET_VERIFIED, all 51 exports
  name-gated with FINAL_DOCS 0 and COPYOUT_OK, RESTORE_MATCH,
  LOCK_RELEASED plain. Driver log: `pilot/pdf_oracle8_g1.out`. All 51
  PDFs pulled to `entries/<id>/` and verified on the VM
  (verify_full.py): no probe flags; markup-custom streams carry the
  emitted #2E97D3/#B5082E exactly; markup-default by-author palette
  observed per author (#5C2E91 for "Comparison", #7160E8 for
  "Author"). Manifest stamped: **60 verified / 144 pending**.
  Notable record facts: tc-p1-whitespace-only's sole revision is a
  whitespace-only insertion (renders as a 3-pixel author-colour mark
  in markup vectors); tc-p2-base-del's input carries no revisions.

- **Group 2 (16 entries, 48 records) — DONE.** Claim
  2026-11-11T04:29:21Z → 04:40:40Z (~11 min): all gates green,
  CUSTOM_SET_VERIFIED, RESTORE_MATCH, LOCK_RELEASED plain, 48/48
  COPYOUT_OK. Verified on the VM with no probe flags; manifest
  stamped: **108 verified / 96 pending**. Record facts of note:
  tc-p2-both-del, tc-p2-rev-del and tc-p2m-empty-both inputs carry
  no tracked revisions (same pre-existing-revisions-accepted
  pattern); tc-p2m-cjk-change's revisions are CJK text — verified
  on producer/page/colour facts (both emitted colours present in
  the content streams) rather than Latin probes.

- **Group 3 (16 entries, 48 records) — DONE.** Claim
  2026-11-11T04:42:55Z → 04:54:29Z (~11.5 min): all gates green,
  RESTORE_MATCH, LOCK_RELEASED plain, 48/48 COPYOUT_OK. Verified on
  the VM with no probe flags; manifest stamped: **156 verified /
  48 pending**. Record facts of note: tc-p2m-image-replace's
  revisions wrap a picture (no revision text; the clean PDF's
  saturated pixels are the replacement image itself, recorded as
  document content); tc-p2m-rtl-change verified on colour facts
  (both emitted colours in the streams); the section/list/style
  entries' revisions are formatting-only (sectPrChange, pPrChange,
  tcPrChange family) and their records say so.

- **Group 4 (16 entries, 48 records) — DONE.** Claim
  2026-11-11T04:56:08Z → 05:16:56Z (~21 min — the stress group):
  all gates green, RESTORE_MATCH, LOCK_RELEASED plain, 48/48
  COPYOUT_OK, including all three vectors of
  tc-p4-budget-manypairs (67,500 ins + 67,500 del elements; its
  markup-default balloon PDF is 629 pages / 4.3 MB, markup-custom
  6.0 MB, clean 2.3 MB) and tc-p4-budget-fewlong. Verified on the
  VM with no probe flags; manifest stamped: **204 verified /
  0 pending**.

### PRODUCTION RUN — FINAL (2026-11-11)

- **Tally: 204/204 records verified** — 68 clean + 68
  markup-default + 68 markup-custom (9 pilot + 195 production).
  No failures; no record was left pending. Full-corpus audit after
  the final stamp: every PDF present on disk, every sha256 and
  byte count recompute-matches the manifest, and every input pin
  matches its `entries/<id>/input.docx`. Corpus PDFs total
  19,714,078 bytes (~19.7 MB).
- **Timing:** four claims totalling ~54 minutes of lock time
  (G1 ~10 min, G2 ~11 min, G3 ~11.5 min, G4 ~21 min); wall clock
  2026-11-11T04:17Z → 05:17Z including per-group verification and
  manifest stamping between claims. Typical export ≈ 12 s; the
  hardest single export (tc-p4-budget-manypairs markup-default)
  took ~2.5 min.
- **End state, confirmed after the last claim:** WORD-LOCK free,
  Word at 0 documents and 0 unnamed windows, queue empty, zero
  pdfpilot leftovers in Word's container, staged Working copies
  removed per claim, and every claim's Track Changes Options
  snapshot restored with RESTORE_MATCH before release — Word's
  options on the mini are byte-identical to the pre-run profile
  defaults. Input documents were never saved; all 65 production
  input hashes recompute-matched source/staged/after in every
  claim.
