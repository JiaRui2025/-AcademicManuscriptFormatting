#!/usr/bin/env python3
"""docx_inventory.py -- read-only inventory of a manuscript's references,
figures and tables, for review and for building the numbering mapping table.

Reports, per reference entry:
  number | bookmark present | DOI link present + target | first 70 chars of entry
and per figure/table:
  bookmark present | caption text

Output is human-readable text, or CSV with --csv (one row per reference entry).
Never modifies the input file.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("ERROR: python-docx is required (pip install python-docx)\n")
    raise SystemExit(2)

W = qn
ENTRY_RE = re.compile(r"^\[(\d+)\]\s*(.*)$")
FIG_CAP_RE = re.compile(r"^(Fig\.|Table)\s*(\d+)\s*\.?\s*(.*)$")
DOI_RE = re.compile(r"(10\.\d{4,9}/[^\s\"<>]+)")


def para_text(p):
    return "".join(t.text or "" for t in p.findall(".//" + W("w:t")))


def collect(path):
    doc = docx.Document(path)
    body = doc.element.body
    paras = body.findall(".//" + W("w:p"))

    names = set()
    for b in body.findall(".//" + W("w:bookmarkStart")):
        n = b.get(W("w:name"))
        if n:
            names.add(n)

    # external DOI hyperlink targets, in document order
    doi_targets = []
    for h in body.findall(".//" + W("w:hyperlink")):
        rid = h.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if not rid:
            continue
        rel = doc.part.rels.get(rid)
        if rel is not None and rel.is_external:
            doi_targets.append((h, rel.target_ref))

    ref_start = len(paras)
    for i, p in enumerate(paras):
        if re.fullmatch(r"(References?|REFERENCES|参考文献|文献)",
                        para_text(p).strip()):
            ref_start = i
            break

    refs = []
    for i in range(ref_start, len(paras)):
        m = ENTRY_RE.match(para_text(paras[i]).strip())
        if not m:
            continue
        num, rest = int(m.group(1)), m.group(2)
        # find a DOI link inside this paragraph
        para_dois = []
        for h, tgt in doi_targets:
            if h.getparent() is not None and _contains(paras[i], h):
                if tgt and "doi.org" in tgt:
                    para_dois.append(tgt)
        refs.append({
            "num": num,
            "bookmark": "_ref_%d" % num in names,
            "doi": para_dois[0] if para_dois else "",
            "inline_doi": bool(DOI_RE.search(rest)),
            "text": rest,
            "para": i,
        })

    figs, tbls = [], []
    for i in range(ref_start):
        m = FIG_CAP_RE.match(para_text(paras[i]).strip())
        if not m:
            continue
        kind, num, cap = m.group(1), int(m.group(2)), m.group(3)
        bm = "_fig_%d" % num if kind.startswith("Fig") else "_tbl_%d" % num
        rec = {"num": num, "bookmark": bm in names, "caption": cap, "para": i}
        (figs if kind.startswith("Fig") else tbls).append(rec)

    # how many in-text links point at each figure bookmark
    link_counts = {}
    for h in body.findall(".//" + W("w:hyperlink")):
        a = h.get(W("w:anchor"))
        if a:
            link_counts[a] = link_counts.get(a, 0) + 1

    return {"file": path, "refs": refs, "figs": figs, "tbls": tbls,
            "link_counts": link_counts, "bookmarks": names}


def _contains(para, el):
    for d in para.iter():
        if d is el:
            return True
    return False


def main():
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--csv", help="write the reference mapping to this CSV file")
    args = ap.parse_args()

    inv = collect(args.docx)
    refs, figs, tbls = inv["refs"], inv["figs"], inv["tbls"]

    print("[references] %d entries" % len(refs))
    print("  %-5s %-9s %-46s %s" % ("no.", "bookmark", "DOI link", "entry"))
    for r in refs:
        doi = r["doi"] if r["doi"] else ("(text only)" if r["inline_doi"] else "-")
        print("  %-5d %-9s %-46s %s"
              % (r["num"], "ok" if r["bookmark"] else "MISSING",
                 doi[:44], r["text"][:60]))

    print("\n[figures] %d" % len(figs))
    for f in figs:
        n = inv["link_counts"].get("_fig_%d" % f["num"], 0)
        print("  Fig. %-3d bookmark=%-5s links=%-3d %s"
              % (f["num"], "ok" if f["bookmark"] else "MISSING",
                 n, f["caption"][:70]))

    print("\n[tables] %d" % len(tbls))
    for t in tbls:
        n = inv["link_counts"].get("_tbl_%d" % t["num"], 0)
        print("  Table %-3d bookmark=%-5s links=%-3d %s"
              % (t["num"], "ok" if t["bookmark"] else "MISSING",
                 n, t["caption"][:70]))

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["new_number", "bookmark", "doi_link", "in_text_doi",
                        "first_cited_para", "entry"])
            for r in refs:
                w.writerow([r["num"], "_ref_%d" % r["num"],
                            r["doi"], "yes" if r["inline_doi"] else "no",
                            r["para"], r["text"]])
        print("\n[csv] written to %s" % args.csv)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
