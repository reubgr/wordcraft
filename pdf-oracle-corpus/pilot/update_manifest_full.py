#!/usr/bin/env python3
"""update_manifest_full.py — stamp production-run records verified (#6153).

Usage: update_manifest_full.py <group> <snapshot-file> <id> [<id> ...]

Reads pilot/evidence_g<group>.json (from verify_full.py) and the claim's
Track Changes Options snapshot line, then stamps the given entries'
three records verified with the same settings-block shape the pilot
used (update_manifest.py): full options vector parsed from the actual
claim snapshot, verification evidence, fonts notes. Asserts: targets
are currently `pending`, their input pins are unchanged, every PDF
hash recomputes against the evidence, and all non-target records are
byte-identical after a JSON round trip.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
MANIFEST = os.path.join(ROOT, "manifest.json")

PRODUCER = "macOS Version 15.8.2 (Build 24H232) Quartz PDFContext"
BUILD = "Word for Mac 16.114.4 (26110224)"

SNAP2OPT = [
    ("insColor", "inserted_text_color"), ("insMark", "inserted_text_mark"),
    ("delColor", "deleted_text_color"), ("delMark", "deleted_text_mark"),
    ("insCellColor", "inserted_cell_color"), ("delCellColor", "deleted_cell_color"),
    ("mergedCellColor", "merged_cell_color"), ("splitCellColor", "split_cell_color"),
    ("moveFromColor", "move_from_text_color"), ("moveFromMark", "move_from_text_mark"),
    ("moveToColor", "move_to_text_color"), ("moveToMark", "move_to_text_mark"),
    ("commentsColor", "comments_color"), ("printComments", "print_comments"),
]

EMITTED_NOTE = ("Word does NOT emit the raw enum RGB in the PDF: with inserted "
                "text color set to `blue` and deleted text color set to `red`, "
                "the PDF content streams (ICCBased colour space, `sc` operators) "
                "carry insertions as 0.1803922 0.5921569 0.8274510 (#2E97D3) and "
                "deletions as 0.7098039 0.03137255 0.1803922 (#B5082E) — "
                "confirmed identically by 150-DPI raster sampling. A scorer for "
                "this vector must expect Word's emitted palette colours, not "
                "0000FF/FF0000.")
BALLOON_NOTE = (" Word for Mac's default markup PDF puts deletions in balloons: the deleted content appears in the "
                "PDF text layer as 'Deleted: <text>' margin annotations. The by-author palette assigns a different "
                "colour per author — the behaviour issue #5232 is about.")


def parse_snapshot(path):
    line = open(path).read().strip()
    fields = {}
    for part in line.replace("\t", " ").split():
        if "=" in part:
            k, v = part.split("=", 1)
            fields[k] = v
    opts = {}
    for sk, mk in SNAP2OPT:
        v = fields.get(sk, "")
        if sk == "printComments" and (not v or v.startswith("ERR") or v == "missing"):
            v = "missing value (unset)"
        opts[mk] = v
    return opts, fields


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
    group = sys.argv[1]
    snapfile = sys.argv[2]
    ids = sys.argv[3:]
    orig_opts, snap_fields = parse_snapshot(snapfile)
    assert orig_opts["inserted_text_color"] and not orig_opts["inserted_text_color"].startswith("ERR"), snap_fields
    custom_opts = dict(orig_opts, inserted_text_color="blue", inserted_text_mark="underline",
                       deleted_text_color="red", deleted_text_mark="strikethrough")
    snap_str = ", ".join(f"{k}={snap_fields[k]}" for k, _ in SNAP2OPT if k in snap_fields)
    readback_orig = (f"COMPLETE (production run 2026-11-10, group {group} claim): equality-based snapshot via the "
                     f"application `settings` property (Word options object, wSet) — {snap_str}. "
                     "Taken at claim start and re-taken identically after the restore step (RESTORE_MATCH in the "
                     "driver log); these are the untouched profile defaults in effect for the markup-default "
                     "and clean exports.")
    readback_custom = ("COMPLETE: custom values set via the application `settings` property and verified by "
                       "equality readback before export (insColor=blue, insMark=underline, delColor=red, "
                       "delMark=strikethrough; all other options at the profile defaults snapshotted the same "
                       "claim). Originals restored after the exports and verified by an identical RESTORED "
                       "snapshot (RESTORE_MATCH).")
    evidence = json.load(open(os.path.join(HERE, f"evidence_g{group}.json")))
    inv = json.load(open(os.path.join(HERE, "inventory.json")))

    with open(MANIFEST) as f:
        m = json.load(f)
    target_ids = set(ids)
    before_others = {json.dumps(r, sort_keys=True) for r in m["records"] if r["id"] not in target_ids}
    pins = {}
    n = 0
    for r in m["records"]:
        if r["id"] not in target_ids:
            continue
        key = f"{r['id']}|{r['vector']}"
        ev = evidence[key]
        assert not ev.get("missing"), f"PDF missing for {key}"
        assert r["oracle_status"] == "pending", f"{key} not pending: {r['oracle_status']}"
        pins[key] = (r["input"]["sha256"], r["input"]["bytes"])
        path = os.path.join(ROOT, "entries", r["id"], f"word-{r['vector']}.pdf")
        import hashlib
        h = hashlib.sha256(open(path, "rb").read()).hexdigest()
        assert h == ev["sha256"] and os.path.getsize(path) == ev["bytes"], f"evidence hash mismatch {key}"
        r["oracle_pdf"] = {"file": f"entries/{r['id']}/word-{r['vector']}.pdf",
                           "sha256": ev["sha256"], "bytes": ev["bytes"]}
        r["oracle_status"] = "verified"
        r["word_build"] = BUILD
        r["fonts_note"] = ev["fonts_note"]
        s = base_settings(r["vector"])
        if ev["page_size"] and "612 x 792" not in ev["page_size"]:
            s["paper_size"] = ev["page_size"]
        if r["vector"] == "clean":
            s.update({
                "markup_shown": False,
                "show_revisions_set": False,
                "print_revisions_set": False,
                "revisions_view": "final",
                "revisions_mode": "balloon (profile view state at export; immaterial with markup hidden)",
                "comments_shown": False,
                "track_changes_options_state": "profile defaults, untouched by the run",
                "track_changes_options": orig_opts,
                "track_changes_options_readback": readback_orig,
            })
        elif r["vector"] == "markup-default":
            s.update({
                "markup_shown": True,
                "show_revisions_set": True,
                "print_revisions_set": True,
                "revisions_view": "final (showing markup)",
                "revisions_mode": "balloon (deletions render as margin balloons — see detail)",
                "comments_shown": True,
                "track_changes_options_state": "profile defaults, untouched by the run",
                "track_changes_options": orig_opts,
                "track_changes_options_readback": readback_orig,
            })
        else:
            s.update({
                "markup_shown": True,
                "show_revisions_set": True,
                "print_revisions_set": True,
                "revisions_view": "final (showing markup)",
                "revisions_mode": "inline (in-line revisions set at export; balloons inline per FORMAT.md)",
                "comments_shown": True,
                "track_changes_options_state": "set explicitly for this vector via the application `settings` property; originals restored and verified after export",
                "track_changes_options": custom_opts,
                "track_changes_options_readback": readback_custom,
                "emitted_revision_colours": {"insertions": "#2E97D3", "deletions": "#B5082E"},
            })
        s["verification"] = {
            "pages": ev["pages"],
            "producer": ev["producer"],
            "evidence": ev["evidence"],
            "verified_on": "2026-11-10 (production run; pilot-report.md)",
        }
        detail = (f"Produced by the production run (pdf_oracle8 driver, group {group}, 2026-11-10) in one "
                  "WORD-LOCK claim (holder pdf-oracle-pilot); name-gated export, single container save + "
                  "shell copy-out.")
        iv = inv[r["id"]]
        if iv["ins_count"] + iv["del_count"] == 0 and not iv["fmt_changes"]:
            detail += (" The input carries no tracked revisions (verified by XML inventory of the .docx), "
                       "so the markup vectors render the document without revision markup.")
        elif iv["fmt_changes"] and not iv["has_text_revisions"]:
            detail += (" The input's revisions are formatting-only (" + ", ".join(iv["fmt_changes"]) +
                       "), so no inserted/deleted text renders as revision text.")
        if r["vector"] == "markup-custom":
            detail += " " + EMITTED_NOTE
        if r["vector"] == "markup-default":
            detail += BALLOON_NOTE
        s["detail"] = detail
        r["settings"] = s
        n += 1
    assert n == 3 * len(ids), n
    after_others = {json.dumps(r, sort_keys=True) for r in m["records"] if r["id"] not in target_ids}
    assert before_others == after_others, "non-target records changed!"
    for r in m["records"]:
        key = f"{r['id']}|{r['vector']}"
        if key in pins:
            assert (r["input"]["sha256"], r["input"]["bytes"]) == pins[key], f"pin changed {key}"
    counts = {}
    for r in m["records"]:
        counts[r["oracle_status"]] = counts.get(r["oracle_status"], 0) + 1
    with open(MANIFEST, "w") as f:
        json.dump(m, f, indent=2)
        f.write("\n")
    print(f"group {group}: stamped {n} records verified; status counts: {counts}")


if __name__ == "__main__":
    main()
