#!/usr/bin/env python3
"""prep_inputs.py — stage the 65 pending entries' inputs on the VM (#6153).

For every manifest entry with pending records: locate the #6045 source
triplet's word-compare.docx, verify sha256+bytes against the manifest
pin, and copy it to entries/<id>/input.docx. Pilot entries (already
verified) are checked but not overwritten. Prints a summary and the
sorted list of pending entry ids.
"""
import hashlib
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
SRC_ROOT = os.path.expanduser("~/workspace/wordcraft-oracle-corpus/triplets")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    with open(os.path.join(ROOT, "manifest.json")) as f:
        m = json.load(f)
    entries = {}
    for r in m["records"]:
        entries.setdefault(r["id"], r)
    pending_ids = sorted({r["id"] for r in m["records"] if r["oracle_status"] == "pending"})
    verified_ids = sorted({r["id"] for r in m["records"] if r["oracle_status"] == "verified"})
    bad = []
    for eid, r in sorted(entries.items()):
        pin_sha, pin_bytes = r["input"]["sha256"], r["input"]["bytes"]
        trip = r["provenance"]["source_triplet"]
        src = os.path.join(SRC_ROOT, trip, r["provenance"]["source_file"])
        dst = os.path.join(ROOT, "entries", eid, "input.docx")
        if not os.path.exists(src):
            bad.append((eid, "SOURCE MISSING " + src))
            continue
        s_sha, s_bytes = sha256(src), os.path.getsize(src)
        if (s_sha, s_bytes) != (pin_sha, pin_bytes):
            bad.append((eid, f"SOURCE PIN MISMATCH sha={s_sha} bytes={s_bytes}"))
            continue
        if eid in pending_ids:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copyfile(src, dst)
            if sha256(dst) != pin_sha:
                bad.append((eid, "STAGED COPY HASH MISMATCH"))
        else:
            if not os.path.exists(dst) or sha256(dst) != pin_sha:
                bad.append((eid, "EXISTING PILOT INPUT MISSING/MISMATCH"))
    print(f"entries total={len(entries)} pending={len(pending_ids)} verified={len(verified_ids)}")
    if bad:
        print("PROBLEMS:")
        for eid, msg in bad:
            print(" ", eid, msg)
        sys.exit(1)
    print("all input pins verify; pending inputs staged under entries/<id>/input.docx")
    print("PENDING_IDS=" + " ".join(pending_ids))


if __name__ == "__main__":
    main()
