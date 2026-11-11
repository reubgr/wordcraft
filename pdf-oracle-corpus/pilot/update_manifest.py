#!/usr/bin/env python3
"""update_manifest.py — stamp the 9 pilot records verified (#6153).

Rewrites ONLY the 9 pilot records (3 entries x clean/markup-default/
markup-custom): oracle_pdf pointers for the 6 newly produced PDFs,
oracle_status=verified, full settings vectors (including the completed
Track Changes Options readback via the application `settings`
property), verification evidence, and fonts notes. Asserts before
saving: the other 195 records are byte-identical after a JSON round
trip, and the pilot records' input pins are unchanged.
"""
import copy
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "..", "manifest.json")

PRODUCER = "macOS Version 15.8.2 (Build 24H232) Quartz PDFContext"
BUILD = "Word for Mac 16.114.4 (26110224)"

ORIGINAL_OPTS = {
    "inserted_text_color": "byAuthor",
    "inserted_text_mark": "underline",
    "deleted_text_color": "byAuthor",
    "deleted_text_mark": "strikethrough",
    "inserted_cell_color": "lightBlue",
    "deleted_cell_color": "pink",
    "merged_cell_color": "lightYellow",
    "split_cell_color": "lightOrange",
    "move_from_text_color": "green",
    "move_from_text_mark": "doubleStrikethrough",
    "move_to_text_color": "green",
    "move_to_text_mark": "doubleUnderline",
    "comments_color": "byAuthor",
    "print_comments": "missing value (unset)",
}
CUSTOM_OPTS = dict(ORIGINAL_OPTS,
                   inserted_text_color="blue",
                   inserted_text_mark="underline",
                   deleted_text_color="red",
                   deleted_text_mark="strikethrough")

READBACK_ORIG = ("COMPLETE (resumption 2026-11-10): equality-based snapshot via the "
                 "application `settings` property (Word options object, wSet) — "
                 "insColor=byAuthor, insMark=underline, delColor=byAuthor, "
                 "delMark=strikethrough, insCellColor=lightBlue, delCellColor=pink, "
                 "mergedCellColor=lightYellow, splitCellColor=lightOrange, "
                 "moveFromColor=green, moveFromMark=doubleStrikethrough, "
                 "moveToColor=green, moveToMark=doubleUnderline, "
                 "commentsColor=byAuthor, printComments=missing value (unset). "
                 "Taken 2026-11-11T02:37Z (probe) and re-taken identically at the "
                 "v7 production run; these are the untouched profile defaults that "
                 "were in effect for the markup-default and clean exports.")
READBACK_CUSTOM = ("COMPLETE: custom values set via the application `settings` "
                   "property and verified by equality readback before export "
                   "(insColor=blue, insMark=underline, delColor=red, "
                   "delMark=strikethrough; all other options at the profile "
                   "defaults snapshotted the same run). Originals restored after "
                   "the exports and verified by an identical RESTORED snapshot.")

EMITTED_NOTE = ("Word does NOT emit the raw enum RGB in the PDF: with inserted "
                "text color set to `blue` and deleted text color set to `red`, "
                "the PDF content streams (ICCBased colour space, `sc` operators) "
                "carry insertions as 0.1803922 0.5921569 0.8274510 (#2E97D3) and "
                "deletions as 0.7098039 0.03137255 0.1803922 (#B5082E) — "
                "confirmed identically by 150-DPI raster sampling. A scorer for "
                "this vector must expect Word's emitted palette colours, not "
                "0000FF/FF0000.")

NEW_PDFS = {
    ("tc-p0-del-para", "clean"): ("854cf7a73148d17ad34ce3e3270ba08da19c122e13878cb27be6e0bd67356a88", 8265),
    ("tc-p0-del-para", "markup-custom"): ("0a81dd7993277fb3e85f2d01894b853207ba91de64d6e03adec278f67981d1f2", 8590),
    ("tc-p2m-multiauthor", "clean"): ("47d8d32ef8e5c8a6f24c4efb75ebb52d908006dd742503752a81706eb46d5b99", 8604),
    ("tc-p2m-multiauthor", "markup-custom"): ("972f8d782f9825e40fa581b80d52ebe332369d7006907d62ec371155d6b1010f", 8731),
    ("tc-p2m-table-row-del", "clean"): ("514bbfce6b5d72e068e6df6825c691f01319b49855143b0ff77b75c40a57d807", 7552),
    ("tc-p2m-table-row-del", "markup-custom"): ("eaf0bf32ba7ce3592b5f95f7ab3db9f3a1870c00cd13c546da46109c0cc125a9", 7986),
}

FONTS = {
    ("tc-p0-del-para", "clean"): "Cambria (embedded subset)",
    ("tc-p0-del-para", "markup-custom"): "Cambria (embedded subset)",
    ("tc-p0-del-para", "markup-default"): "Cambria, TimesNewRomanPS-BoldMT (embedded subsets)",
    ("tc-p2m-multiauthor", "clean"): "Cambria (embedded subset)",
    ("tc-p2m-multiauthor", "markup-custom"): "Cambria (embedded subset)",
    ("tc-p2m-multiauthor", "markup-default"): "Cambria (embedded subset)",
    ("tc-p2m-table-row-del", "clean"): "Cambria (embedded subset)",
    ("tc-p2m-table-row-del", "markup-custom"): "Cambria (embedded subset)",
    ("tc-p2m-table-row-del", "markup-default"): "Cambria, Cambria-Bold, TimesNewRomanPS-BoldMT, TimesNewRomanPSMT (embedded subsets)",
}

PAGES = {
    ("tc-p0-del-para", "clean"): 1, ("tc-p0-del-para", "markup-default"): 1, ("tc-p0-del-para", "markup-custom"): 1,
    ("tc-p2m-multiauthor", "clean"): 1, ("tc-p2m-multiauthor", "markup-default"): 1, ("tc-p2m-multiauthor", "markup-custom"): 1,
    ("tc-p2m-table-row-del", "clean"): 1, ("tc-p2m-table-row-del", "markup-default"): 2, ("tc-p2m-table-row-del", "markup-custom"): 1,
}

EVIDENCE = {
    ("tc-p0-del-para", "clean"): ("Text (pdftotext): deleted paragraph absent — 'staffing of the night desk' not present; "
                                  "no 'Deleted:' annotation. Colour sample at 150 DPI: zero saturated pixels (no revision colours)."),
    ("tc-p0-del-para", "markup-custom"): ("Text: deleted paragraph present INLINE (no balloon): 'staffing of the night desk' in the page text, "
                                          "no 'Deleted:' prefix — balloons-inline (in-line revisions view) confirmed. Colour: deletions emitted "
                                          "#B5082E (content-stream `sc` 0.7098039 0.03137255 0.1803922; raster top #B5082E)."),
    ("tc-p2m-multiauthor", "clean"): ("Text: insertions present in final text ('extended timetable', 'final draft'); no 'Deleted:'. "
                                      "Colour sample at 150 DPI: zero saturated pixels."),
    ("tc-p2m-multiauthor", "markup-custom"): ("Text: insertions present ('extended timetable', 'final draft'). Colour: insertions emitted "
                                              "#2E97D3 (content-stream `sc` 0.1803922 0.5921569 0.8274510; raster top #2E97D3, 142 samples). "
                                              "Entry has no deleted text, so no deletion colour applies."),
    ("tc-p2m-table-row-del", "clean"): ("Text: deleted row content absent ('middle-one' not present); no 'Deleted:'. Colour sample: zero "
                                        "saturated pixels. 1 page (the markup-default balloon layout takes 2)."),
    ("tc-p2m-table-row-del", "markup-custom"): ("Text: deleted row cells present INLINE ('middle-one'), no 'Deleted:' balloon prefix. Colour: "
                                                "deletions emitted #B5082E (content-stream `sc`; raster top #B5082E). 1 page."),
    ("tc-p0-del-para", "markup-default"): ("Run-3 evidence (2026-11-09): deleted paragraph present as balloon annotation "
                                           "'Deleted: Attention then turned to the staffing of the night desk during the winter holiday "
                                           "period.' (re-confirmed 2026-11-10 by plain pdftotext). Colour sample at 150 DPI: red-dominant "
                                           "by-author palette (top #D13438)."),
    ("tc-p2m-multiauthor", "markup-default"): ("Run-3 evidence (2026-11-09): insertions present in final text ('extended timetable', "
                                               "'final draft'). Colour sample at 150 DPI: blue-dominant by-author palette (content-stream "
                                               "`sc` 0 0.4705882 0.8313726 = #0078D4)."),
    ("tc-p2m-table-row-del", "markup-default"): ("Run-3 evidence (2026-11-09): deleted row cells present as balloon annotation "
                                                 "'Deleted: middle-one ...' with author/date attribution. Colour sample at 150 DPI: "
                                                 "purple-dominant by-author palette (top #5C2E91, 1,716 samples)."),
}

PROD_NOTE = {
    ("tc-p0-del-para", "markup-custom"): ("Produced by the v6 driver export 1 (2026-11-11T02:43Z) — the only valid export of that run: "
                                          "DOC_NAME verified as this entry's staged copy, custom settings readback-verified active, "
                                          "container save + shell copy-out. (The run's later exports were invalidated and redone "
                                          "by v7 after the Grant File Access wedge was cleared; see pilot-report.md.)"),
    ("tc-p2m-multiauthor", "markup-custom"): "Produced by the v7 driver (2026-11-11T03:18Z) after the session was cleared; name-gated export, container save + shell copy-out.",
    ("tc-p2m-table-row-del", "markup-custom"): "Produced by the v7 driver (2026-11-11T03:18Z) after the session was cleared; name-gated export, container save + shell copy-out.",
    ("tc-p0-del-para", "clean"): "Produced by the v7 driver (2026-11-11T03:19Z); name-gated export, container save + shell copy-out.",
    ("tc-p2m-multiauthor", "clean"): "Produced by the v7 driver (2026-11-11T03:19Z); name-gated export, container save + shell copy-out.",
    ("tc-p2m-table-row-del", "clean"): "Produced by the v7 driver (2026-11-11T03:19Z); name-gated export, container save + shell copy-out.",
    ("tc-p0-del-para", "markup-default"): "Produced by pilot run 3 (2026-11-09); settings vector completed at the 2026-11-10 resumption.",
    ("tc-p2m-multiauthor", "markup-default"): "Produced by pilot run 3 (2026-11-09); settings vector completed at the 2026-11-10 resumption.",
    ("tc-p2m-table-row-del", "markup-default"): "Produced by pilot run 3 (2026-11-09); settings vector completed at the 2026-11-10 resumption.",
}


def base_settings(vector):
    return {
        "vector": vector,
        "route": "word-applescript-saveas-pdf",
        "format_string": "format PDF (AppleScript dictionary save as)",
        "fields_updated_before_export": False,
        "paper_size": "letter (612 x 792 pts)",
        "producer": PRODUCER,
    }


def main():
    with open(MANIFEST) as f:
        m = json.load(f)
    before_others = {json.dumps(r, sort_keys=True) for r in m["records"]
                     if r["id"] not in ("tc-p0-del-para", "tc-p2m-multiauthor", "tc-p2m-table-row-del")}
    pilot_pins = {}
    n = 0
    for r in m["records"]:
        key = (r["id"], r["vector"])
        if r["id"] not in ("tc-p0-del-para", "tc-p2m-multiauthor", "tc-p2m-table-row-del"):
            continue
        n += 1
        pilot_pins[key] = (r["input"]["sha256"], r["input"]["bytes"])
        if key in NEW_PDFS:
            sha, b = NEW_PDFS[key]
            r["oracle_pdf"] = {"file": f"entries/{r['id']}/word-{r['vector']}.pdf",
                               "sha256": sha, "bytes": b}
        r["oracle_status"] = "verified"
        r["word_build"] = BUILD
        r["fonts_note"] = FONTS[key]
        s = base_settings(r["vector"])
        if r["vector"] == "clean":
            s.update({
                "markup_shown": False,
                "show_revisions_set": False,
                "print_revisions_set": False,
                "revisions_view": "final",
                "revisions_mode": "balloon (profile view state at export; immaterial with markup hidden)",
                "comments_shown": False,
                "track_changes_options_state": "profile defaults, untouched by the pilot",
                "track_changes_options": ORIGINAL_OPTS,
                "track_changes_options_readback": READBACK_ORIG,
            })
        elif r["vector"] == "markup-default":
            s.update({
                "markup_shown": True,
                "show_revisions_set": True,
                "print_revisions_set": True,
                "revisions_view": "final (showing markup)",
                "revisions_mode": "balloon (deletions render as margin balloons — see detail)",
                "comments_shown": True,
                "track_changes_options_state": "profile defaults, untouched by the pilot",
                "track_changes_options": ORIGINAL_OPTS,
                "track_changes_options_readback": READBACK_ORIG,
            })
        else:  # markup-custom
            s.update({
                "markup_shown": True,
                "show_revisions_set": True,
                "print_revisions_set": True,
                "revisions_view": "final (showing markup)",
                "revisions_mode": "inline (in-line revisions set at export; balloons inline per FORMAT.md)",
                "comments_shown": True,
                "track_changes_options_state": "set explicitly for this vector via the application `settings` property; originals restored and verified after export",
                "track_changes_options": CUSTOM_OPTS,
                "track_changes_options_readback": READBACK_CUSTOM,
                "emitted_revision_colours": {"insertions": "#2E97D3", "deletions": "#B5082E"},
            })
        s["verification"] = {
            "pages": PAGES[key],
            "producer": PRODUCER,
            "evidence": EVIDENCE[key],
            "verified_on": "2026-11-10 (resumption run; pilot-report.md)",
        }
        detail = PROD_NOTE[key]
        if r["vector"] == "markup-custom":
            detail += " " + EMITTED_NOTE
        if r["vector"] == "markup-default":
            detail += (" Word for Mac's default markup PDF puts deletions in balloons: the deleted content appears in the "
                       "PDF text layer as 'Deleted: <text>' margin annotations. The by-author palette assigns a different "
                       "colour per author — the behaviour issue #5232 is about.")
        s["detail"] = detail
        r["settings"] = s
    assert n == 9, n
    after_others = {json.dumps(r, sort_keys=True) for r in m["records"]
                    if r["id"] not in ("tc-p0-del-para", "tc-p2m-multiauthor", "tc-p2m-table-row-del")}
    assert before_others == after_others, "non-pilot records changed!"
    for r in m["records"]:
        key = (r["id"], r["vector"])
        if key in pilot_pins:
            assert (r["input"]["sha256"], r["input"]["bytes"]) == pilot_pins[key], f"pin changed {key}"
    counts = {}
    for r in m["records"]:
        counts[r["oracle_status"]] = counts.get(r["oracle_status"], 0) + 1
    with open(MANIFEST, "w") as f:
        json.dump(m, f, indent=2)
        f.write("\n")
    print("status counts:", counts)
    print("pilot records verified:", sum(1 for r in m["records"] if r["oracle_status"] == "verified"))


if __name__ == "__main__":
    main()
