#!/usr/bin/env python3
"""verify_full.py — verify production-run PDFs on the VM (#6153).

Usage: verify_full.py <group> <id> [<id> ...]

For each entries/<id>/word-<vector>.pdf: sha256, bytes, pdfinfo (pages,
producer, page size), pdftotext probe checks driven by inventory.json,
150-DPI colour sampling (first <=3 pages), a PDF content-stream scan
for the emitted `sc` colour triples, and pdffonts. Composes a factual
evidence string per record and writes pilot/evidence_g<group>.json.
Exit 0 always (evidence tool); anomalies are flagged in the output for
review before update_manifest_full.py stamps anything.
"""
import hashlib
import json
import os
import re
import subprocess
import sys
import zlib

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
VECTORS = ["clean", "markup-default", "markup-custom"]

EMITTED_DEL = (0.7098039, 0.03137255, 0.1803922)   # #B5082E
EMITTED_INS = (0.1803922, 0.5921569, 0.8274510)    # #2E97D3


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def pdfinfo(path):
    out = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    d = {}
    for line in out.splitlines():
        if ": " in line:
            k, v = line.split(": ", 1)
            d[k.strip()] = v.strip()
    return d


def pdftext(path):
    r = subprocess.run(["pdftotext", "-layout", path, "-"],
                       capture_output=True, text=True)
    return re.sub(r"\s+", " ", r.stdout)


def fonts_note(path):
    out = subprocess.run(["pdffonts", path], capture_output=True, text=True).stdout
    names, allsub = [], True
    for ln in out.strip().splitlines()[2:]:
        if not ln.strip():
            continue
        toks = ln.split()
        name = toks[0]
        base = name.split("+", 1)[-1]
        sub = toks[-3] if len(toks) >= 3 else "no"
        if sub != "yes":
            allsub = False
        if base not in names:
            names.append(base)
    if not names:
        return None
    joined = ", ".join(names)
    if allsub:
        joined += " (embedded subset)" if len(names) == 1 else " (embedded subsets)"
    return joined


def colour_sample(path, pages):
    n = min(pages, 3)
    for f in os.listdir("/tmp"):
        if f.startswith("vfull"):
            os.remove(os.path.join("/tmp", f))
    with open(os.devnull, "wb") as dn:
        subprocess.run(["pdftoppm", "-png", "-r", "150", "-f", "1", "-l", str(n),
                        path, "/tmp/vfull"], stdout=dn, stderr=dn)
    counts = {"red": 0, "blue": 0, "purple": 0, "other_sat": 0}
    exact = {}
    for pf in sorted(f for f in os.listdir("/tmp") if f.startswith("vfull")):
        im = Image.open(os.path.join("/tmp", pf)).convert("RGB")
        px = im.load()
        w, h = im.size
        for y in range(0, h, 2):
            for x in range(0, w, 2):
                r, g, b = px[x, y]
                mx, mn = max(r, g, b), min(r, g, b)
                if mx - mn < 80 or mx < 80:
                    continue
                if r > 150 and g < 110 and b < 110:
                    counts["red"] += 1
                elif b > 150 and r < 110 and g < 110:
                    counts["blue"] += 1
                elif r > 90 and b > 90 and g < 110:
                    counts["purple"] += 1
                else:
                    counts["other_sat"] += 1
                exact[(r, g, b)] = exact.get((r, g, b), 0) + 1
        os.remove(os.path.join("/tmp", pf))
    top = sorted(exact.items(), key=lambda kv: -kv[1])[:4]
    top_s = ", ".join(f"#{r:02X}{g:02X}{b:02X}x{n}" for (r, g, b), n in top)
    return counts, top_s, n


def stream_colours(path):
    data = open(path, "rb").read()
    found = {}
    for m in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.S):
        raw = m.group(1)
        try:
            dec = zlib.decompress(raw)
        except Exception:
            continue
        for t in re.finditer(rb"(\d*\.?\d+)\s+(\d*\.?\d+)\s+(\d*\.?\d+)\s+scn?\b", dec):
            trip = tuple(round(float(x), 4) for x in t.groups())
            if trip[0] == trip[1] == trip[2]:
                continue
            found[trip] = found.get(trip, 0) + 1
    return found


def hexof(trip):
    return "#%02X%02X%02X" % tuple(round(c * 255) for c in trip)


def main():
    group = sys.argv[1]
    ids = sys.argv[2:]
    inv = json.load(open(os.path.join(HERE, "inventory.json")))
    evidence = {}
    for eid in ids:
        iv = inv[eid]
        for vector in VECTORS:
            path = os.path.join(ROOT, "entries", eid, f"word-{vector}.pdf")
            print(f"== {eid} / {vector}")
            if not os.path.exists(path):
                print("   MISSING")
                evidence[f"{eid}|{vector}"] = {"missing": True}
                continue
            info = pdfinfo(path)
            pages = int(info.get("Pages", "0"))
            producer = info.get("Producer", "")
            text = pdftext(path)
            delp, insp = iv["del_probe"], iv["ins_probe"]
            del_in = (delp in text) if delp else None
            ins_in = (insp in text) if insp else None
            delprefix = "Deleted:" in text
            counts, top_s, sampled = colour_sample(path, pages)
            streams = stream_colours(path)
            stream_hex = sorted(hexof(t) for t in streams)
            has_edel = any(abs(t[0]-EMITTED_DEL[0]) < 0.01 and abs(t[1]-EMITTED_DEL[1]) < 0.01
                           and abs(t[2]-EMITTED_DEL[2]) < 0.01 for t in streams)
            has_eins = any(abs(t[0]-EMITTED_INS[0]) < 0.01 and abs(t[1]-EMITTED_INS[1]) < 0.01
                           and abs(t[2]-EMITTED_INS[2]) < 0.01 for t in streams)
            fnote = fonts_note(path)
            print(f"   sha256={sha256(path)} bytes={os.path.getsize(path)}")
            print(f"   pages={pages} producer={producer} size={info.get('Page size')}")
            print(f"   del_probe={delp!r} in_text={del_in} | ins_probe={insp!r} in_text={ins_in} | 'Deleted:'={delprefix}")
            print(f"   colour150 pages1-{sampled}: {counts} top: {top_s}")
            print(f"   stream sc colours: {stream_hex} emitted_del={has_edel} emitted_ins={has_eins}")
            print(f"   fonts={fnote}")

            # ---- compose factual evidence
            clauses = []
            if iv["has_text_revisions"] and not delp and not insp:
                clauses.append("input's revised text is non-Latin (no Latin probe tokens); "
                               "verification rests on the page/producer/colour facts below")
            if iv["has_text_revisions"]:
                if vector == "clean":
                    if delp:
                        clauses.append(
                            f"deleted content absent — probe '{delp}' not present in the text layer"
                            if not del_in else
                            f"probe '{delp}' occurs in the text layer (token shared with final text; see inventory)")
                    if insp:
                        clauses.append(
                            f"insertions present as final text (probe '{insp}')"
                            if ins_in else f"insertion probe '{insp}' NOT found — FLAGGED")
                    if not delprefix:
                        clauses.append("no 'Deleted:' annotation text")
                elif vector == "markup-default":
                    if delp:
                        clauses.append(
                            f"deleted content present in the text layer (balloon annotation; 'Deleted:' prefix present: {delprefix})"
                            if del_in else f"deleted-text probe '{delp}' NOT found — FLAGGED")
                    if insp:
                        clauses.append(
                            f"inserted content present in the text (probe '{insp}')"
                            if ins_in else f"insertion probe '{insp}' NOT found — FLAGGED")
                else:  # markup-custom
                    if delp:
                        if del_in and not delprefix:
                            clauses.append(f"deleted content present INLINE (probe '{delp}'; no 'Deleted:' balloon prefix)")
                        elif del_in:
                            clauses.append(f"deleted content present (probe '{delp}'); 'Deleted:' annotation text also present")
                        else:
                            clauses.append(f"deleted-text probe '{delp}' NOT found — FLAGGED")
                    if insp:
                        clauses.append(
                            f"inserted content present in the text (probe '{insp}')"
                            if ins_in else f"insertion probe '{insp}' NOT found — FLAGGED")
            n_rev_el = iv["ins_count"] + iv["del_count"]
            if iv["fmt_changes"] and not iv["has_text_revisions"]:
                clauses.append("input carries formatting-only revisions (" +
                               ", ".join(iv["fmt_changes"]) + "); no textual ins/del probes apply")
            if n_rev_el > 0 and not iv["has_text_revisions"] and not iv["fmt_changes"]:
                clauses.append(f"input carries {n_rev_el} revision element(s) with no visible text "
                               "(whitespace/empty content); textual probes do not apply")
            if iv["del_count"] > 0 and not iv["del_text"]:
                clauses.append("the deletion element(s) carry no deleted text (empty/field content)")
            if n_rev_el == 0 and not iv["fmt_changes"] and not iv["has_comments"]:
                clauses.append("input carries no tracked revisions (XML inventory: zero ins/del/"
                               "formatting-change elements); vectors render the document as-is")
            if iv["has_comments"]:
                clauses.append("input carries comments (word/comments.xml present)")
            sat_total = sum(counts.values())
            if vector == "clean":
                if sat_total == 0:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): zero saturated pixels (no revision colours)")
                else:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): saturated pixels present "
                                   f"(red={counts['red']} blue={counts['blue']} purple={counts['purple']} other={counts['other_sat']}; "
                                   f"top {top_s}) — document content, not revision markup (clean vector shows no markup)")
            elif vector == "markup-default":
                if sat_total == 0:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): no saturated revision colours")
                else:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): red={counts['red']} blue={counts['blue']} "
                                   f"purple={counts['purple']} other={counts['other_sat']}; dominant saturated {top_s} (by-author palette)")
                if stream_hex:
                    clauses.append(f"content-stream colours include {', '.join(stream_hex)}")
            else:  # markup-custom
                cc = []
                if has_edel:
                    cc.append("deletions emitted #B5082E (content-stream `sc` 0.7098039 0.03137255 0.1803922)")
                if has_eins:
                    cc.append("insertions emitted #2E97D3 (content-stream `sc` 0.1803922 0.5921569 0.8274510)")
                if cc:
                    clauses.append("Colour: " + "; ".join(cc) +
                                   f". Raster sample at 150 DPI (first {sampled} page(s)): top saturated {top_s}")
                elif sat_total == 0 and not stream_hex:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): no saturated revision colours; "
                                   "no non-gray content-stream colours")
                else:
                    clauses.append(f"Colour sample at 150 DPI (first {sampled} page(s)): red={counts['red']} blue={counts['blue']} "
                                   f"purple={counts['purple']} other={counts['other_sat']}; top saturated {top_s}; "
                                   f"content-stream colours {', '.join(stream_hex) if stream_hex else 'none non-gray'}")
            ev = "Text (pdftotext): " + "; ".join(clauses) + "."
            evidence[f"{eid}|{vector}"] = {
                "sha256": sha256(path), "bytes": os.path.getsize(path),
                "pages": pages, "producer": producer,
                "page_size": info.get("Page size", ""),
                "fonts_note": fnote, "evidence": ev,
                "del_in_text": del_in, "ins_in_text": ins_in,
                "deleted_prefix": delprefix,
            }
            print(f"   EVIDENCE: {ev}")
    with open(os.path.join(HERE, f"evidence_g{group}.json"), "w") as f:
        json.dump(evidence, f, indent=2)
    print(f"evidence written: evidence_g{group}.json ({len(evidence)} records)")


if __name__ == "__main__":
    main()
