#!/usr/bin/env python3
"""
Relink duplicate Mendeley Cite citations in a .docx so one paper stops
rendering as 2021a / 2021b.

Mendeley Cite stores each citation's library-item ID inside the document.
If two library records of the same paper were cited (and one was later
deleted), the document still holds two IDs, so APA adds a/b. This script
rewrites every citation of a duplicated paper to a single ID.

Usage
  python fix_mendeley_dupes.py paper.docx                # dry run: list duplicate groups
  python fix_mendeley_dupes.py paper.docx --apply        # write paper.fixed.docx
  python fix_mendeley_dupes.py paper.docx --apply --only 10.1257/pol.20190414
  python fix_mendeley_dupes.py paper.docx --apply --keep <item-id>

Which ID is kept (per duplicate group):
  --keep ID  if given; otherwise the ID of the LAST citation of that paper
  in the document. Tip: before running, insert one fresh citation of the
  paper at the very end of the document from your library. That makes the
  surviving library record the keeper. Delete that extra citation afterwards.

The original file is never modified.
"""
import argparse
import base64
import copy
import html
import json
import re
import sys
import zipfile
from collections import defaultdict

PREFIX = "MENDELEY_CITATION_v3_"
BLOB_RE = re.compile(re.escape(PREFIX) + r"([A-Za-z0-9+/]+={0,2})")
TAG_RE = re.compile(r'<w:tag\s+w:val="' + re.escape(PREFIX) + r'([A-Za-z0-9+/]+={0,2})"')
ATTR_JSON_RE = re.compile(r'value="(\[[^"]*\]|\{[^"]*\})"')
CITATION_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")


# ---------- base64 <-> JSON (round-trip safe) ----------

def decode_blob(b64):
    """Return (obj, encode_fn) or (None, None) if not decodable."""
    padded = b64 + "=" * (-len(b64) % 4)
    raw = base64.b64decode(padded)
    for codec in ("utf-8", "latin-1"):
        try:
            txt = raw.decode(codec)
            obj = json.loads(txt)
        except (UnicodeDecodeError, ValueError):
            continue
        keep_pad = b64.endswith("=") or len(b64) % 4 == 0

        def encode(o, codec=codec, keep_pad=keep_pad):
            s = json.dumps(o, ensure_ascii=False, separators=(",", ":"))
            out = base64.b64encode(s.encode(codec)).decode("ascii")
            return out if keep_pad else out.rstrip("=")

        return obj, encode
    return None, None


# ---------- scanning ----------

def norm_doi(d):
    d = (d or "").strip().lower()
    return re.sub(r"^(https?://(dx\.)?doi\.org/|doi:)", "", d)


def norm_title(t):
    return re.sub(r"[^a-z0-9]+", "", (t or "").lower())


def group_key(item_data):
    doi = norm_doi(item_data.get("DOI"))
    if doi:
        return "doi:" + doi
    year = ""
    try:
        year = str(item_data["issued"]["date-parts"][0][0])
    except (KeyError, IndexError, TypeError):
        pass
    fam = ""
    try:
        fam = norm_title(item_data["author"][0].get("family"))
    except (KeyError, IndexError, TypeError, AttributeError):
        pass
    return f"t:{norm_title(item_data.get('title'))}|{year}|{fam}"


def scan(zf):
    """Collect every cited item in document order."""
    occ = []  # (seq, id, itemData, citeprocText)
    seq = 0
    for part in CITATION_PARTS:
        if part not in zf.namelist():
            continue
        xml = zf.read(part).decode("utf-8")
        for m in TAG_RE.finditer(xml):
            obj, _ = decode_blob(m.group(1))
            if not isinstance(obj, dict):
                continue
            seq += 1
            text = (obj.get("manualOverride") or {}).get("citeprocText", "")
            for it in obj.get("citationItems", []):
                iid = it.get("id") or (it.get("itemData") or {}).get("id")
                if iid:
                    occ.append((seq, str(iid), it.get("itemData") or {}, text))
    return occ, seq


def find_groups(occ, only=None, keep_ids=()):
    by_key = defaultdict(lambda: defaultdict(list))
    for seq, iid, data, text in occ:
        by_key[group_key(data)][iid].append((seq, data, text))
    groups = []
    for key, ids in by_key.items():
        if len(ids) < 2:
            continue
        sample = next(iter(ids.values()))[0][1]
        if only:
            hay = (norm_doi(sample.get("DOI")) + " " + (sample.get("title") or "")).lower()
            if not any(o.lower() in hay or o in ids for o in only):
                continue
        forced = [i for i in ids if i in keep_ids]
        keeper = forced[0] if forced else max(ids, key=lambda i: max(s for s, _, _ in ids[i]))
        groups.append({"key": key, "ids": ids, "keeper": keeper, "sample": sample})
    return groups


# ---------- rewriting ----------

class Remapper:
    def __init__(self, remap, canon_data):
        self.remap = remap            # old_id -> keeper_id
        self.canon = canon_data       # keeper_id -> itemData

    def walk(self, o):
        """Mutate o in place. Returns True if anything changed."""
        changed = False
        if isinstance(o, dict):
            items = o.get("citationItems")
            if isinstance(items, list):
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    old = str(it.get("id") or (it.get("itemData") or {}).get("id") or "")
                    if old in self.remap:
                        new = self.remap[old]
                        it["id"] = new
                        if isinstance(it.get("itemData"), dict) and new in self.canon:
                            it["itemData"] = copy.deepcopy(self.canon[new])
                        changed = True
            for k, v in list(o.items()):
                if isinstance(v, (dict, list)):
                    changed |= self.walk(v)
                elif isinstance(v, str):
                    nv = self.fix_str(v)
                    if nv != v:
                        o[k] = nv
                        changed = True
        elif isinstance(o, list):
            for i, v in enumerate(o):
                if isinstance(v, (dict, list)):
                    changed |= self.walk(v)
                elif isinstance(v, str):
                    nv = self.fix_str(v)
                    if nv != v:
                        o[i] = nv
                        changed = True
        return changed

    def fix_str(self, s):
        if s.startswith(PREFIX):
            return PREFIX + self.fix_b64(s[len(PREFIX):])
        for old, new in self.remap.items():
            if old in s:
                s = s.replace(old, new)
        return s

    def fix_b64(self, b64):
        obj, enc = decode_blob(b64)
        if obj is None or not self.walk(obj):
            return b64  # untouched -> byte-identical
        return enc(obj)

    def fix_xml(self, xml):
        n = 0

        # 1) JSON stored in XML attributes (Mendeley Cite's document settings)
        def attr_sub(m):
            nonlocal n
            raw = m.group(1)
            try:
                obj = json.loads(html.unescape(raw))
            except ValueError:
                return m.group(0)
            if not self.walk(obj):
                return m.group(0)
            n += 1
            s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
            s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            return f'value="{s}"'

        xml = ATTR_JSON_RE.sub(attr_sub, xml)

        # 2) base64 citation blobs (content-control tags)
        def blob_sub(m):
            nonlocal n
            new = self.fix_b64(m.group(1))
            if new != m.group(1):
                n += 1
            return PREFIX + new

        xml = BLOB_RE.sub(blob_sub, xml)

        # 3) any plain-text leftovers (e.g., legacy Mendeley Desktop fields)
        for old, new in self.remap.items():
            c = xml.count(old)
            if c:
                n += c
                xml = xml.replace(old, new)
        return xml, n


def rewrite(src, dst, remapper):
    total = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith((".xml", ".rels")):
                try:
                    xml = data.decode("utf-8")
                except UnicodeDecodeError:
                    xml = None
                if xml is not None:
                    new_xml, n = remapper.fix_xml(xml)
                    if n:
                        total += n
                        print(f"  {info.filename}: {n} change(s)")
                        data = new_xml.encode("utf-8")
            zout.writestr(info, data)
    return total


# ---------- CLI ----------

def label(data):
    try:
        who = data["author"][0].get("family", "?")
    except (KeyError, IndexError, TypeError, AttributeError):
        who = "?"
    try:
        yr = data["issued"]["date-parts"][0][0]
    except (KeyError, IndexError, TypeError):
        yr = "n.d."
    title = (data.get("title") or "")[:70]
    return f"{who} ({yr}) — {title}"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--apply", action="store_true", help="write <name>.fixed.docx")
    ap.add_argument("--only", action="append", help="limit to groups whose DOI/title/ID contains this (repeatable)")
    ap.add_argument("--keep", action="append", default=[], help="force this item ID as keeper (repeatable)")
    ap.add_argument("-o", "--out", help="output path (default: <name>.fixed.docx)")
    a = ap.parse_args()

    with zipfile.ZipFile(a.docx) as zf:
        occ, ncit = scan(zf)
        legacy = any(
            b"CSL_CITATION" in zf.read(p) for p in CITATION_PARTS if p in zf.namelist()
        )

    print(f"Mendeley Cite citations found: {ncit}  (cited items: {len(occ)})")
    if legacy:
        print("Note: legacy Mendeley Desktop fields detected; they get best-effort ID replacement only.")
    if not ncit and not legacy:
        sys.exit("No Mendeley citations found in this document.")

    groups = find_groups(occ, a.only, set(a.keep))
    if not groups:
        print("No duplicate groups found (no paper is cited under more than one ID).")
        return

    print(f"\nDuplicate groups: {len(groups)}\n")
    remap, canon = {}, {}
    for i, g in enumerate(groups, 1):
        print(f"[{i}] {label(g['sample'])}")
        print(f"    key: {g['key']}")
        for iid, hits in sorted(g["ids"].items(), key=lambda kv: -len(kv[1])):
            texts = sorted({t for _, _, t in hits if t})[:2]
            mark = "  <- KEEP" if iid == g["keeper"] else ""
            print(f"    id {iid}  cited {len(hits):>4}x  last at citation #{max(s for s, _, _ in hits)}"
                  f"  e.g. {', '.join(texts) or '-'}{mark}")
            if iid != g["keeper"]:
                remap[iid] = g["keeper"]
        canon[g["keeper"]] = g["ids"][g["keeper"]][-1][1]
        print()

    if not a.apply:
        print("Dry run only. Re-run with --apply to write the fixed copy.")
        return

    out = a.out or re.sub(r"\.docx$", "", a.docx, flags=re.I) + ".fixed.docx"
    print(f"Writing {out}")
    n = rewrite(a.docx, out, Remapper(remap, canon))
    print(f"Done: {n} change(s). Open it in Word -> Mendeley Cite -> More -> Refresh References.")


if __name__ == "__main__":
    main()
