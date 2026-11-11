#!/usr/bin/env python3
"""verify_pdfs.py — verify pilot PDFs on the VM (#6153).

For each entries/<id>/word-<vector>.pdf present: sha256, bytes, pdfinfo
(pages, producer, page size), pdftotext content checks per entry/vector,
and rendered colour sampling at 150 DPI (pdftoppm + PIL): counts of
strongly saturated red / blue / purple pixels, used to check the
markup-custom vector's exact colours (blue 0000FF ins, red FF0000 del)
and the markup-default by-author palette.

Prints one evidence block per PDF. Exit 0 always (evidence tool).
"""
import hashlib
import subprocess
import sys
import os

from PIL import Image

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "entries")
IDS = ["tc-p0-del-para", "tc-p2m-multiauthor", "tc-p2m-table-row-del"]
VECTORS = ["clean", "markup-default", "markup-custom"]

# (must_contain, must_not_contain) per (entry, vector-class)
CHECKS = {
    "tc-p0-del-para": {
        "markup": (["Deleted:", "staffing of the night desk"], []),
        "clean": ([], ["staffing of the night desk", "Deleted:"]),
    },
    "tc-p2m-multiauthor": {
        "markup": (["extended timetable", "final draft"], []),
        "clean": (["extended timetable", "final draft"], ["Deleted:"]),
    },
    "tc-p2m-table-row-del": {
        "markup": (["Deleted:", "middle-one"], []),
        "clean": ([], ["middle-one", "Deleted:"]),
    },
}


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
    return subprocess.run(
        ["pdftotext", "-layout", path, "-"], capture_output=True, text=True
    ).stdout


def colour_sample(path):
    """Render page 1 (and page 2 if present) at 150 DPI; count saturated
    red/blue/purple pixels and report the dominant exact colours."""
    with open(os.devnull, "wb") as dn:
        subprocess.run(["pdftoppm", "-png", "-r", "150", path, "/tmp/vpdf"],
                       stdout=dn, stderr=dn)
    pages = sorted(f for f in os.listdir("/tmp") if f.startswith("vpdf"))
    counts = {"red": 0, "blue": 0, "purple": 0, "other_sat": 0}
    exact = {}
    for pf in pages:
        im = Image.open(os.path.join("/tmp", pf)).convert("RGB")
        px = im.load()
        w, h = im.size
        step = 2  # sample every 2nd pixel for speed
        for y in range(0, h, step):
            for x in range(0, w, step):
                r, g, b = px[x, y]
                mx, mn = max(r, g, b), min(r, g, b)
                if mx - mn < 80 or mx < 80:  # not saturated
                    continue
                if r > 150 and g < 110 and b < 110:
                    counts["red"] += 1
                elif b > 150 and r < 110 and g < 110:
                    counts["blue"] += 1
                elif r > 90 and b > 90 and g < 110:
                    counts["purple"] += 1
                else:
                    counts["other_sat"] += 1
                key = (r, g, b)
                exact[key] = exact.get(key, 0) + 1
        os.remove(os.path.join("/tmp", pf))
    top = sorted(exact.items(), key=lambda kv: -kv[1])[:4]
    top_s = ", ".join(f"#{r:02X}{g:02X}{b:02X}x{n}" for (r, g, b), n in top)
    return counts, top_s


def main():
    for entry in IDS:
        for vector in VECTORS:
            path = os.path.join(ROOT, entry, f"word-{vector}.pdf")
            print(f"== {entry} / {vector}")
            if not os.path.exists(path):
                print("   MISSING")
                continue
            info = pdfinfo(path)
            print(f"   sha256={sha256(path)} bytes={os.path.getsize(path)}")
            print(f"   pages={info.get('Pages')} producer={info.get('Producer')} "
                  f"page_size={info.get('Page size')}")
            text = pdftext(path)
            cls = "clean" if vector == "clean" else "markup"
            must, mustnot = CHECKS[entry][cls]
            for s in must:
                print(f"   TEXT contains {s!r}: {s in text}")
            for s in mustnot:
                print(f"   TEXT lacks {s!r}: {s not in text}")
            fonts = subprocess.run(["pdffonts", path], capture_output=True,
                                   text=True).stdout.strip().splitlines()
            fnames = sorted({ln.split()[0] for ln in fonts[2:] if ln.strip()})
            print(f"   fonts={', '.join(fnames)}")
            counts, top = colour_sample(path)
            print(f"   colour150 red={counts['red']} blue={counts['blue']} "
                  f"purple={counts['purple']} other_sat={counts['other_sat']}")
            print(f"   top saturated colours: {top}")


if __name__ == "__main__":
    main()
