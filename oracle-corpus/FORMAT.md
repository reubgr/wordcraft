# Word-Oracle Compare Corpus — Format

Seed package for storytold/wordcraft#6045: a project-owned regression
corpus of Word-oracle compare triplets. The corpus exists because a
compare engine's own tests can only verify that the engine does what
it intends — never that it does what Word does. Every entry here is a
document pair plus the output that a real, build-stamped Word
produced from that pair under a fully recorded settings vector, plus
a ground-truth edit script for the pair.

## Entry layout

```
triplets/<id>/
  base.docx          the original document
  revised.docx       the revised document
  word-compare.docx  Word's own compare output for the pair
  truth.json         the ground-truth edit script for the pair
manifest.json        one record per entry (see schema below)
```

Entry ids are stable and prefixed by origin (`p0-`, `p1-`, `p2-`,
`p2m-`, `p3-`, `p4-` in this seed, from the study phases that produced
them). An id, once published, never changes and is never reused.

## Manifest schema

`manifest.json` is `{corpus, issue, generated_from, entry_count,
entries[]}`. Each entry record:

| Field | Meaning |
|---|---|
| `id` | Entry id; matches the `triplets/<id>/` directory name. |
| `source_phase` | Which study phase produced the pair (`phase0` … `phase4`). |
| `files.base` / `files.revised` / `files.word_compare` / `files.truth` | `{file, sha256, bytes}` for each of the four files in the entry directory. The sha256 is the integrity pin: any consumer must verify a file against the manifest before trusting it. |
| `word_build` | The exact Word build that produced `word-compare.docx` (e.g. `Word for Mac 16.114.4 (26110224)`). `null` when the producing phase's records do not stamp a build — never inferred, never backfilled from another phase. An entry without a build stamp is usable for scoring but must be re-baselined before it anchors a release claim. |
| `wordcraft_head` | The pinned WordCraft commit whose compare output was scored against this oracle in the source study (provenance for the original measurements; the oracle itself is Word's). `null` when the phase did not record a pin. |
| `settings` | The compare settings vector under which the oracle was produced (see convention below). `{scope: "entry" \| "phase", word: {...}, wordcraft: {...}}` — the Word vector is the oracle's; the WordCraft vector records what the study ran against it. |
| `truth_summary` | `{change_type, description}` lifted from the entry's `truth.json` (the full edit script — token streams, substitutions — is in `truth.json` itself). |
| `provenance` | `{study_corpus_dir, word_compare_source_file}`: where the entry came from and which source file became `word-compare.docx`. |
| `notes` | Present only when the entry is non-routine (e.g. Word produced its result through a different entry point, or the on-disk canonical filename in the source corpus was not a real Word output). Read it before scoring the entry. |

## Settings-vector convention

**Every field explicit; no defaults are ever implied.** A settings
vector records every compare toggle (insertions/deletions, moves,
comments, formatting, case changes, white space, tables,
headers/footers, footnotes/endnotes, textboxes, fields), the
granularity level (`word` or `character`), the destination (`new`,
`original`, or `revised` document), and the author/label the run
stamped its revisions with. A run whose vector is not fully recorded
is not a corpus run — it cannot be scored against an entry, because
any divergence might be a settings difference rather than an engine
difference. (One historical exception is documented in the seed:
Phase 0 ran the Word arm at captured out-of-box defaults, before this
convention existed; its phase-level settings record says exactly
that.)

## Word build stamp requirement

Every new entry must record the full Word build that produced its
oracle (product, version, and build number where the platform exposes
one, e.g. `Word for Mac 16.114.4 (26110224)`). Word's compare
behaviour can change between builds; an unstamped oracle silently
expires. When Word updates, entries are not edited in place — the
re-baselining cadence (which entries get re-run, and whether old
oracles are kept as a prior-build tier) is an open design question on
issue #6045.

## Scoring guidance

Score a compare engine against an entry **three ways, and report all
three**:

1. **Vs ground truth** — the engine's revisions against the edit
   script in `truth.json` (inserted/deleted token streams, ins/del
   F1). This measures whether the engine found the real changes.
2. **Vs the Word oracle** — the engine's revisions against
   `word-compare.docx`. This measures fidelity to Word's behaviour.
3. **Accept/Reject invariants** — accepting all revisions in the
   engine's output must reproduce the revised document exactly, and
   rejecting all must reproduce the base document exactly (token
   streams, and part-level state where the change type touches it).

Truth and Word can disagree: Word itself sometimes marks more or
less than the true edit, and sometimes deviates from its own
documented semantics. Neither score substitutes for the other, and a
divergence between them is a finding to report, not noise to average
away. Token counting must count `w:t` inside `w:del` as deleted text
(the encoding Word and WordCraft both emit); a scorer that reads
only `w:delText` understates deletions and will not reproduce any
published baseline.

## Adding an entry

1. Source documents must be **constructed or public-domain seeds
   only** — never real client or user documents — so the corpus stays
   freely redistributable.
2. Produce the oracle on a real Word, at a fully explicit settings
   vector, and save the result as a new document.
3. Write `truth.json`: `id`, `change_type`, `description`, and the
   token-level edit script (`inserted_tokens`, `deleted_tokens`,
   `substituted` pairs) for the constructed change.
4. Create `triplets/<id>/` with the four files, compute sha256 +
   byte size for each, and append the manifest record (all fields;
   `word_build` mandatory for new entries).
5. Verify: every sha256 in the manifest recomputes against the files
   on disk, and every `.docx` opens as a valid zip.

## What this seed deliberately excludes

Settings-variant runs. The source study ran every pair under many
vectors (one-toggle-off arms, character level, alternate
destinations, control arms); only the **baseline-vector triplet**
(base, revised, Word at the baseline vector, truth) is in the seed,
one per pair. Variant runs are the obvious future extension — as
additional manifest records keyed to the same pair id with their own
full vectors — but they are not mixed into the seed, so that "the
oracle" for a pair always means one unambiguous document.
