# Word-Oracle PDF Corpus — Format

Seed package for storytold/wordcraft#6153: a project-owned regression
corpus of Word-oracle PDF pairs. The corpus exists because a PDF
exporter's own tests can only verify that the exporter does what it
intends — never that it renders what Word renders. Every entry here
is a document plus the PDF that a real, build-stamped Word produced
from that document under fully recorded export settings vectors,
plus a feature tag list that makes coverage a manifest query.

## Entry layout

```
entries/<id>/
  input.docx            the document exported
  word-clean.pdf        Word's PDF under the clean vector
  word-markup-default.pdf
                        Word's PDF under the markup-default vector
  word-markup-custom.pdf
                        Word's PDF under the markup-custom vector
                        (tracked entries only, where applicable)
  word-balloons.pdf     Word's PDF under the balloons vector
                        (entries with comments/formatting, where applicable)
manifest.json           one record per (entry, vector) — see schema below
```

Entry ids are stable and prefixed by arm: `tc-` (tracked arm, seeded
from the #6045 compare corpus), `mx-` (constructed layout/rendering
matrix), `wild-` (neurotic-bench wild arm). An id, once published,
never changes and is never reused.

## The four canonical vectors

A Word PDF is a function of its settings, so an entry is exported
under several vectors; each vector is its own manifest record.

| Vector | Meaning |
|---|---|
| `clean` | Document only — final text, tracked deletions omitted (Word's No Markup equivalent). |
| `markup-default` | Document showing markup, Word's Track Changes Options untouched. |
| `markup-custom` | Document showing markup, Track Changes Options set explicitly: insertions blue `0000FF` underline, deletions red `FF0000` strikethrough, balloons inline. This is the #5232 vector. |
| `balloons` | Document showing markup with comments/formatting in balloons beside the scaled page (Print Markup layout). **Not seeded** (owner ruling, 2026-11-11): revision balloons are one rendering choice among many and are not scored; formatting-change balloon annotations (`Formatted: …`) are already captured in the markup-default arm, and comments render in both markup arms. |

**Every field explicit; no defaults are ever implied.** A vector
record states: markup shown or not; the full Track Changes Options
state (insert/delete colours and marks, move colours and marks,
formatting, balloons mode, comment display); comments shown or not;
the export route used in Word (Save As PDF vs Print to PDF — they
are not assumed equivalent; the route is recorded, and any entry
produced by both carries both records); paper size handling; and
whether fields were updated before export. A run whose vector is not
fully recorded is not a corpus run.

## Manifest schema

`manifest.json` is `{corpus, issue, generated_from, record_count,
records[]}`. Each record:

| Field | Meaning |
|---|---|
| `id` | Entry id; matches the `entries/<id>/` directory name. |
| `arm` | `tracked` \| `matrix` \| `wild`. |
| `vector` | One of the four canonical vectors (or a recorded extension vector). |
| `input` | `{file, sha256, bytes}` for `input.docx`. The sha256 is the integrity pin. |
| `oracle_pdf` | `{file, sha256, bytes}` for the Word PDF, or `null` while `oracle_status` is `pending`. |
| `oracle_status` | `pending` \| `produced` \| `verified` (verified = PDF opens, page count and producer recorded, text extractable). |
| `word_build` | The exact Word build that produced the oracle PDF (e.g. `Word for Mac 16.114.4 (26110224)`). Mandatory for produced records; `null` while pending — never inferred. |
| `settings` | The full export settings vector (see convention above). |
| `feature_tags` | Tags from the controlled vocabulary below. |
| `provenance` | Where the input came from (e.g. `{source_corpus: "wordcraft-oracle-corpus (#6045)", source_triplet: "p0-del-para", source_file: "word-compare.docx"}`). |
| `fonts_note` | Recorded font facts for the entry (faces the document names, faces Word's PDF embeds) — font substitution is a recorded property, not silent noise. `null` until produced. |
| `notes` | Present only when the record is non-routine. |

## Feature tag vocabulary (v1)

Layout/rendering: `tables`, `tables-nested`, `tables-floating`,
`lists-multilevel`, `sections-multiple`, `columns`, `headers-footers`,
`footnotes`, `endnotes`, `fields`, `toc`, `pictures-anchored`,
`shapes`, `textboxes`, `charts`, `smartart`, `equations`,
`hyperlinks-outline`, `content-controls`, `drop-caps`, `watermark`,
`rtl`, `cjk`, `line-numbers`, `page-borders`.

Tracked changes: `tc-ins-del`, `tc-paragraph-marks`, `tc-moves`,
`tc-formatting-run`, `tc-formatting-paragraph`, `tc-formatting-table`,
`tc-formatting-row`, `tc-formatting-cell`, `tc-formatting-section`,
`tc-numbering-change`, `tc-table-rows`, `tc-table-cells`,
`tc-cell-merge`, `tc-in-headers-footers`, `tc-in-footnotes`,
`tc-in-endnotes`, `tc-in-textboxes`, `tc-multiauthor`,
`tc-preexisting` (revisions not produced by WordCraft compare),
`comments`.

The vocabulary is versioned with this file; new tags are added here
first, never ad hoc in a record.

## Scoring guidance

Never compare PDF bytes: Word stamps dates and IDs into its PDFs,
and WordCraft's own export is not byte-reproducible (#3083). Score a
PDF exporter against a record by:

1. **Extracted text** — content and reading order vs the oracle PDF's
   extraction (and, for clean-vector records, vs the document's final
   text).
2. **Page count** — exact match expected; a mismatch is a headline
   finding, reported before any per-page score.
3. **Per-page pixel difference** — rendered pages compared with a
   stated tolerance and rasterisation DPI (the metric the ODT corpus
   work uses), reported per page, worst page first.
4. **Markup colour sampling** (markup vectors) — sampled revision
   colours vs the vector's stated colours; by-author records score
   ins/del distinguishability, custom records score exact colours.
   **Emitted colours, not setting values** (measured 2026-11-10,
   #6153 pilot): Word does not emit the Track Changes Options' raw
   enum RGB. With inserted text colour set to `blue` and deleted
   text colour set to `red`, Word for Mac's PDFs carry insertions
   as **#2E97D3** and deletions as **#B5082E** (ICCBased content
   streams, confirmed by 150-DPI raster sampling) — not `0000FF` /
   `FF0000`. A scorer for the markup-custom vector must expect the
   emitted palette values, which are recorded per record as
   `emitted_revision_colours` in the record's settings.
5. **Structure** — link annotations present, heading outline present
   and free of deleted text (clean) / inclusive of it (markup).

Truth and Word can disagree (fonts Word embeds that the exporter
cannot use are the common case); a divergence is a finding to
report, not noise to average away.

## Adding an entry

1. Source documents must be **constructed or public-domain seeds
   only** — never real client or user documents.
2. Produce each vector's PDF on a real Word at a fully explicit
   settings vector; record the export route.
3. Create `entries/<id>/`, compute sha256 + byte size for the input
   and each PDF, and append one manifest record per vector (all
   fields; `word_build` mandatory once produced).
4. Verify: every sha256 recomputes, every PDF opens and extracts.

## Seed status

The tracked arm's 68 seed records are generated from the #6045
corpus manifest (inputs are its `word-compare.docx` files — real
Word-produced tracked-change documents). All seed records start
`oracle_status: pending`; the first oracle batch produces the
`clean`, `markup-default`, and `markup-custom` PDFs for a pilot
subset, verifies the production route (including how Track Changes
Options colours are set in Word for Mac and which export route
honours them), and only then runs the full set.
