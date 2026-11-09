# Build Report — wordcraft-oracle-corpus (seed for storytold/wordcraft#6045)

Built on this VM from the compare-vs-Word study corpora
(`~/workspace/wordcraft-compare-vs-word/corpus/` + `results/`).
Package: `manifest.json`, `FORMAT.md`, `triplets/` (68 entries,
272 files, 12,274,318 bytes total).

## Entries included, per phase

| Phase | Entries | Word oracle source |
|---|---|---|
| phase0 | 7 | `word.docx` |
| phase1 | 10 | `word-base.docx` (baseline arm) |
| phase2-matrix | 29 | `word.docx` (28) + `word-dlg.docx` (1, see below) |
| phase2-tracked | 11 | `word.docx` (5) + `word-diag-verb.docx` (6, see below) |
| phase3 | 1 | `word-wl.docx` (word-level arm, baseline vector) |
| phase4 | 10 | `word-word-r1.docx` (arm=word, level=word, rep=1) |
| **total** | **68** | |

Every entry carries a per-entry settings record lifted from its
phase's run log (`scope: "entry"`) except Phase 0, whose settings
are phase-level (`scope: "phase"`) for the reason under Skips/notes.

## Non-routine oracle sources (each flagged in its manifest `notes`)

- **p2m-multiauthor** (phase2-matrix): the `word.docx` in the study
  corpus directory is **not** a Word output — its sha256 is identical
  to that directory's `revised.docx` (a copyback placeholder; the VBA
  compare entry point fails deterministically on this pair). The
  genuine oracle, `word-dlg.docx` (Phase 3 dialog-arm repair run at
  the baseline vector; sha256 `5f566711…`, matching the Phase 3
  report), is used as `word-compare.docx`.
- **Six phase2-tracked pairs** (p2-base-del, p2-base-fmt, p2-both-del,
  p2-both-fmt, p2-rev-del, p2-rev-fmt): Word's dialog route produced
  **no result document** (RESULT_NONE — under the tracked-changes
  prompt's Yes semantics the accept-resolved inputs are identical).
  The study's measured Word output for these cells is the verb-route
  equivalent `word-diag-verb.docx` (Word compare verb, warnings
  ignored = prompt Yes, baseline vector), recorded as such in
  `results/phase2-tracked.jsonl`; it is used as `word-compare.docx`.
  No oracle was fabricated: each of these files is a real Word
  compare output for its pair.
- **p3-churn-long** (phase3): only the word-level arm
  (`word-wl.docx`, full baseline vector) enters the seed; the
  character-level arm is a settings variant.
- **Phase 1**: baseline-arm outputs are `word-base.docx`; the
  `word-base-opt.docx` / `word-base-opt2.docx` /
  `word-base-oldtimed.docx` files in two directories are later
  harness re-runs and were not used.
- **Phase 4**: canonical output per `results/word_phase4_runs.jsonl`
  is rep 1 of the word arm at word level; reps r2/r3 are repeats and
  the dense pairs' character-level runs are variants.

## Pairs / runs skipped, and why

- **All settings-variant runs in every pair directory** — OFAT
  one-toggle-off arms (`*-ofat-*`, `word-fnoff/fmtoff/mvoff/tbloff`,
  `word-ctl-*`), character-level arms (`*-char*`), alternate
  destinations (`*-showin-*`, `*-spd-*`), Phase 1 control arms, and
  Phase 4 reps r2/r3. The seed is baseline-vector triplets only, one
  oracle per pair (see FORMAT.md, "What this seed deliberately
  excludes").
- **Phase 3's remaining cells** ran on pairs owned by earlier phases
  (p2m-*, p2-*) under variant settings — no new pairs, no new
  entries; those pairs enter the seed once, from their home phase at
  the baseline vector. (Cell 1, the p2m-multiauthor dialog repair,
  is the one Phase 3 output that does enter, as that pair's missing
  baseline oracle.)
- **Phase 4's re-runs of p0-sub-word, p0-ins-sentence, p0-ins-para,
  p2m-manyblocks, p2m-longpara** were speed arms on existing pairs —
  not new entries.
- **Phase 0's settings caveat (not a skip):** Phase 0 predates the
  study's settings-parity rule. Its Word arm ran at captured
  out-of-box defaults via the AppleScript compare verb (dialog state
  captured that run: all 11 toggles ON, Word level, New document),
  and its WordCraft arm ran CLI defaults + `showIn: new`. The seven
  triplets are genuine Word oracles, but their settings record is
  phase-level and descriptive rather than per-entry and
  set-and-verified; the manifest says so on the phase's settings
  record and FORMAT.md documents the exception.
- No scored baseline pair was dropped for lack of a Word output:
  with the verb-route and dialog-arm outputs above counted, every
  baseline pair in the study corpora has a genuine Word oracle.

## Missing fields (recorded as `null`, never guessed)

- **`word_build` is null for all 29 phase2-matrix entries and for
  p3-churn-long (30 entries).** Neither `phase2-matrix-report.md` /
  `phase3-report.md` nor their jsonl run logs stamp a Word build.
  (The study-wide design doc stamps the study's Word as 16.114.4
  (26110224), and phases 0/1/2-tracked/4 record "Word for Mac
  16.114.4" — with the (26110224) build number in phases 0 and 1 —
  in their own records; the nulls were left null per the
  no-inference rule and should be re-baselined or stamped from the
  study's records before these entries anchor a release claim.)
- **`wordcraft_head` is null for all 10 phase1 entries.** Phase 1's
  report records only "the pinned build" without naming the commit;
  no phase-1 record found on disk names it. All other phases record
  `ab0020d6` (phase4 records the full
  `ab0020d6f4c0b823fa1617b467e1b33b634134db`).

## Verification (run after the build, against the copies)

- All **272 manifest sha256 values recomputed** against the copied
  files: **all match**; all byte sizes match.
- All **204 `.docx` files open as valid zips** (`zipfile` open +
  `testzip` + `[Content_Types].xml` present): **no failures**.
- All **68 `truth.json` files parse**, and each one's internal `id`
  equals its entry id: **no mismatches**.
- Counts per phase: phase0 7, phase1 10, phase2-matrix 29,
  phase2-tracked 11, phase3 1, phase4 10 — **68 total**, matching
  `manifest.json`'s `entry_count`.
