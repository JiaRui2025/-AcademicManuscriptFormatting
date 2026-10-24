#!/usr/bin/env python3
"""format_audit.py -- read-only audit of a .docx against the manuscript house style.

Checks
  1. bookmark hygiene   : every bookmark starts with "_" (Word treats these as hidden)
  2. figure citations   : every "Fig. N" / "Table N" in body text is inside a hyperlink
                          pointing at "_fig_N" / "_tbl_N"
  3. link text purity   : a figure hyperlink's visible text is exactly "Fig. N"
                          (a trailing sub-letter must live OUTSIDE the link)
  4. suffix letters     : sub-letters A..E after a figure link are plain runs --
                          no color 1F4E79, no hyperlink membership
  5. citation brackets  : every "[n]" group is inside a hyperlink to "_ref_n"
  6. anchor integrity   : no w:hyperlink anchor without a matching bookmark
  7. external DOI links : count of r:id hyperlinks whose target is https://doi.org/...

Exit code 0 when the document is clean, 1 when issues were found.
Never modifies the input file.
"""
from __future__ import annotations

import argparse
import json
import re
import sys

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("ERROR: python-docx is required (pip install python-docx)\n")
    raise SystemExit(2)

W = qn
LINK_COLOR = "1F4E79"
FIG_RE = re.compile(r"Fig\.\s*(\d+)")
TBL_RE = re.compile(r"Table\s*(\d+)")
CIT_RE = re.compile(r"\[(\d+(?:\s*[,–\-]\s*\d+)*)\]")


def para_text(p):
    return "".join(t.text or "" for t in p.findall(".//" + W("w:t")))


def run_text(r):
    return "".join(t.text or "" for t in r.findall(".//" + W("w:t")))


def has_link_color(r):
    rpr = r.find(W("w:rPr"))
    if rpr is None:
        return False
    col = rpr.find(W("w:color"))
    return col is not None and col.get(W("w:val")) == LINK_COLOR


def audit(path):
    doc = docx.Document(path)
    body = doc.element.body

    issues = {
        "bookmarks_not_hidden": [],
        "figure_citation_not_linked": [],
        "table_citation_not_linked": [],
        "link_text_not_pure": [],
        "suffix_letter_styled": [],
        "citation_not_linked": [],
        "dangling_anchors": [],
        "missing_fig_bookmarks": [],
    }

    starts = body.findall(".//" + W("w:bookmarkStart"))
    names = set()
    for b in starts:
        n = b.get(W("w:name"))
        if not n:
            continue
        names.add(n)
        if not n.startswith("_"):
            issues["bookmarks_not_hidden"].append(n)

    paras = body.findall(".//" + W("w:p"))

    # The reference list starts at the "References" heading; entry labels there
    # ("[1]", "[2]" ...) are NOT in-text citations and must stay unlinked.
    ref_start = len(paras)
    for i, p in enumerate(paras):
        t = para_text(p).strip()
        if re.fullmatch(r"(References?|REFERENCES|参考文献|文献)", t):
            ref_start = i
            break

    # --- figure / table citations and link-text purity -----------------------
    for i, p in enumerate(paras):
        if i >= ref_start:
            continue
        full = para_text(p)

        # A caption paragraph starts with the label ("Fig. 1.", "Table 2.").
        # Captions are labels, not citations -- they must NOT be hyperlinked.
        caption = re.match(r"^(Fig\.|Table)\s*\d+\s*\.", full) is not None

        # hyperlinks present in this paragraph, in document order
        links = []
        for child in p:
            tag = child.tag.split("}")[1]
            if tag == "hyperlink":
                anc = child.get(W("w:anchor"))
                txt = "".join(t.text or "" for t in child.findall(".//" + W("w:t")))
                links.append((anc, txt))

        # figure citations that must be linked
        if not caption:
            for m in FIG_RE.finditer(full):
                n = int(m.group(1))
                want = "_fig_%d" % n
                if not any(a == want for a, _ in links):
                    issues["figure_citation_not_linked"].append(
                        {"para": i, "fig": n,
                         "context": full[max(0, m.start() - 25):m.end() + 10]}
                    )
            for m in TBL_RE.finditer(full):
                n = int(m.group(1))
                want = "_tbl_%d" % n
                if not any(a == want for a, _ in links):
                    issues["table_citation_not_linked"].append(
                        {"para": i, "table": n,
                         "context": full[max(0, m.start() - 25):m.end() + 10]}
                    )

        # purity: figure/table link text must be exactly "Fig. N" / "Table N"
        for anc, txt in links:
            if not anc:
                continue
            m = re.fullmatch(r"(Fig\.|Table)\s*(\d+)(.*)", txt, re.S)
            if m and m.group(3):
                issues["link_text_not_pure"].append(
                    {"para": i, "anchor": anc, "text": txt,
                     "trailing": m.group(3)}
                )

        # citation brackets must be linked
        for m in CIT_RE.finditer(full):
            first = re.split(r"\s*[,–\-]\s*", m.group(1))[0]
            want = "_ref_%s" % first
            if not any(a == want for a, _ in links):
                issues["citation_not_linked"].append(
                    {"para": i, "cite": m.group(0),
                     "context": full[max(0, m.start() - 20):m.end() + 5]}
                )

        # suffix letters: plain run immediately after a figure/table link
        children = list(p)
        for idx, child in enumerate(children):
            tag = child.tag.split("}")[1]
            if tag != "r":
                continue
            txt = run_text(child)
            if not txt or not txt[0].isalpha() or len(txt) > 2:
                continue
            # look back for the nearest non-empty sibling
            j = idx - 1
            while j >= 0:
                prev_tag = children[j].tag.split("}")[1]
                if prev_tag == "hyperlink":
                    anc = children[j].get(W("w:anchor"))
                    if anc and (anc.startswith("_fig") or anc.startswith("_tbl")):
                        if has_link_color(child):
                            issues["suffix_letter_styled"].append(
                                {"para": i, "letter": txt, "after_anchor": anc}
                            )
                    break
                if prev_tag == "r" and run_text(children[j]).strip():
                    break
                j -= 1

    # --- anchor integrity ----------------------------------------------------
    for h in body.findall(".//" + W("w:hyperlink")):
        a = h.get(W("w:anchor"))
        if a and a not in names:
            issues["dangling_anchors"].append(a)

    # --- expected figure bookmarks ------------------------------------------
    figs = sorted({int(m.group(1)) for p in paras
                   for m in FIG_RE.finditer(para_text(p))})
    for n in figs:
        if "_fig_%d" % n not in names:
            issues["missing_fig_bookmarks"].append("_fig_%d" % n)

    # --- external DOI links ---------------------------------------------------
    rels = doc.part.rels
    doi_links = 0
    for rid, rel in rels.items():
        if rel.is_external and "doi.org" in (rel.target_ref or ""):
            doi_links += 1

    total = sum(len(v) for v in issues.values())
    return {
        "file": path,
        "paragraphs": len(paras),
        "bookmarks": len(names),
        "figure_links": sum(
            1 for h in body.findall(".//" + W("w:hyperlink"))
            if (h.get(W("w:anchor")) or "").startswith("_fig")
        ),
        "citation_links": sum(
            1 for h in body.findall(".//" + W("w:hyperlink"))
            if (h.get(W("w:anchor")) or "").startswith("_ref")
        ),
        "external_doi_links": doi_links,
        "issues": issues,
        "total_issues": total,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", help="path to the .docx to audit (never modified)")
    ap.add_argument("-o", "--out", help="write JSON report here (default: stdout)")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="only print the summary line")
    args = ap.parse_args()

    report = audit(args.docx)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)

    r = report
    print("[summary] %s" % args.docx)
    print("  paragraphs=%(paragraphs)d bookmarks=%(bookmarks)d "
          "fig_links=%(figure_links)d citation_links=%(citation_links)d "
          "doi_links=%(external_doi_links)d" % r)
    print("  total issues: %d" % r["total_issues"])
    if not args.quiet:
        for key, items in r["issues"].items():
            if items:
                print("  - %s (%d)" % (key, len(items)))
                for it in items[:12]:
                    print("      %s" % json.dumps(it, ensure_ascii=False))
                if len(items) > 12:
                    print("      ... and %d more" % (len(items) - 12))
    if args.out:
        print("  JSON report written to %s" % args.out)

    return 1 if r["total_issues"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
