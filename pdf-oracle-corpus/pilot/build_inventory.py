#!/usr/bin/env python3
"""build_inventory.py — revision inventory of pending inputs (#6153 full run).

For each manifest record still `pending`, parse entries/<id>/input.docx
(all story parts) and record: w:ins / w:del counts, inserted & deleted
text, authors, formatting-change elements, comments presence, and
distinctive probe tokens for the verifier. Also asserts every input
sha256 recomputes against the manifest pin. Writes inventory.json.
"""
import hashlib
import json
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
MANIFEST = os.path.join(ROOT, "manifest.json")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def probe_token(text):
    words = re.findall(r"[A-Za-z][A-Za-z0-9'’-]{2,}", text or "")
    if not words:
        return None
    words.sort(key=len, reverse=True)
    for w in words:
        if len(w) >= 5:
            return w
    return words[0]


def inventory(docx_path):
    import xml.etree.ElementTree as ET
    inv = {
        "ins_count": 0, "del_count": 0, "ins_text": "", "del_text": "",
        "authors": [], "fmt_changes": [], "has_comments": False,
        "parts_with_revisions": [],
    }
    authors = set()
    fmt = set()
    with zipfile.ZipFile(docx_path) as z:
        names = [n for n in z.namelist()
                 if n.startswith("word/") and n.endswith(".xml")]
        for name in names:
            data = z.read(name)
            if b"<w:ins" not in data and b"<w:del" not in data \
               and b"Change" not in data and b"cellIns" not in data \
               and b"cellDel" not in data and b"cellMerge" not in data:
                continue
            root = ET.fromstring(data)
            part_has = False
            for el in root.iter():
                tag = el.tag.replace(W, "w:")
                if tag == "w:ins":
                    inv["ins_count"] += 1
                    part_has = True
                    a = el.get(W + "author")
                    if a:
                        authors.add(a)
                    t = "".join(n.text or "" for n in el.iter(W + "t"))
                    inv["ins_text"] += t + " "
                elif tag == "w:del":
                    inv["del_count"] += 1
                    part_has = True
                    a = el.get(W + "author")
                    if a:
                        authors.add(a)
                    t = "".join(n.text or "" for n in el.iter(W + "delText"))
                    inv["del_text"] += t + " "
                elif tag in ("w:moveFrom", "w:moveFromRangeStart"):
                    inv["del_count"] += 1
                    part_has = True
                    a = el.get(W + "author")
                    if a:
                        authors.add(a)
                    t = "".join(n.text or "" for n in el.iter(W + "t"))
                    inv["del_text"] += t + " "
                elif tag in ("w:moveTo", "w:moveToRangeStart"):
                    inv["ins_count"] += 1
                    part_has = True
                    a = el.get(W + "author")
                    if a:
                        authors.add(a)
                    t = "".join(n.text or "" for n in el.iter(W + "t"))
                    inv["ins_text"] += t + " "
                elif tag.endswith("Change") or tag in ("w:cellIns", "w:cellDel", "w:cellMerge"):
                    fmt.add(tag)
                    part_has = True
                    a = el.get(W + "author")
                    if a:
                        authors.add(a)
            if part_has:
                inv["parts_with_revisions"].append(name)
        if "word/comments.xml" in z.namelist():
            inv["has_comments"] = True
    inv["authors"] = sorted(authors)
    inv["fmt_changes"] = sorted(fmt)
    inv["ins_text"] = re.sub(r"\s+", " ", inv["ins_text"]).strip()
    inv["del_text"] = re.sub(r"\s+", " ", inv["del_text"]).strip()
    inv["ins_probe"] = probe_token(inv["ins_text"])
    inv["del_probe"] = probe_token(inv["del_text"])
    inv["has_text_revisions"] = bool(inv["ins_text"] or inv["del_text"])
    return inv


def main():
    with open(MANIFEST) as f:
        m = json.load(f)
    pending_ids = sorted({r["id"] for r in m["records"]
                          if r["oracle_status"] == "pending"})
    pins = {r["id"]: r["input"] for r in m["records"]}
    out = {}
    bad = []
    for eid in pending_ids:
        path = os.path.join(ROOT, "entries", eid, "input.docx")
        actual = sha256(path)
        pin = pins[eid]
        ok = (actual == pin["sha256"] and os.path.getsize(path) == pin["bytes"])
        if not ok:
            bad.append(eid)
        inv = inventory(path)
        inv["pin_ok"] = ok
        out[eid] = inv
        print(f"{eid}: pin_ok={ok} ins={inv['ins_count']} del={inv['del_count']} "
              f"fmt={len(inv['fmt_changes'])} authors={inv['authors']} "
              f"ins_probe={inv['ins_probe']!r} del_probe={inv['del_probe']!r} "
              f"comments={inv['has_comments']}")
    assert not bad, f"pin mismatches: {bad}"
    with open(os.path.join(HERE, "inventory.json"), "w") as f:
        json.dump(out, f, indent=2)
    print(f"inventory written for {len(out)} entries; all pins OK")


if __name__ == "__main__":
    main()
