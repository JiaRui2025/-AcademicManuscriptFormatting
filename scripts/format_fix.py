#!/usr/bin/env python3
"""format_fix.py -- repair a .docx to the manuscript house style.

Operations (each can be enabled independently; all are idempotent):
  --hide-bookmarks    prefix every bookmark name with "_" (Word hides these)
                      and sync every w:hyperlink w:anchor to the new name.
  --fix-suffix-color  strip the link colour (1F4E79) from runs that sit OUTSIDE
                      a hyperlink. These are figure/table sub-letters (A, B, C)
                      that merely inherited the link colour and therefore look
                      clickable but are not.
  --split-link-text   if a figure/table hyperlink's visible text carries a
                      trailing sub-letter ("Fig. 4A"), keep "Fig. 4" inside the
                      link and move the remainder out as a plain run after it.

The input file is never modified; output goes to --out (default: <name>_fixed.docx).
Run format_audit.py afterwards to confirm the result is clean.
"""
from __future__ import annotations

import argparse
import copy
import os
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
LINKISH = re.compile(r"^(Fig\.|Table)\s*(\d+)\s*(.*)$", re.S)


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


def make_plain_run(template_run, text):
    """Build a new run carrying template_run's paragraph-visible formatting
    but with any direct character formatting removed (so it inherits body style)."""
    r = copy.deepcopy(template_run)
    # drop run properties entirely -> inherits style, exactly like body text
    rpr = r.find(W("w:rPr"))
    if rpr is not None:
        r.remove(rpr)
    for t in r.findall(".//" + W("w:t")):
        r.remove(t)
    for br in r.findall(".//" + W("w:br")) + r.findall(".//" + W("w:tab")):
        r.remove(br)
    t = r.makeelement(W("w:t"), {})
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def fix(doc, do_bookmarks, do_suffix, do_split, log):
    body = doc.element.body

    # ---------- 1. hidden bookmarks -----------------------------------------
    if do_bookmarks:
        mapping = {}
        for b in body.findall(".//" + W("w:bookmarkStart")):
            name = b.get(W("w:name"))
            if not name or name.startswith("_"):
                continue
            new = "_" + name
            mapping[name] = new
            b.set(W("w:name"), new)
        synced = 0
        for h in body.findall(".//" + W("w:hyperlink")):
            a = h.get(W("w:anchor"))
            if a and a in mapping:
                h.set(W("w:anchor"), mapping[a])
                synced += 1
        log.append("bookmarks hidden: %d renamed, %d anchors synced"
                   % (len(mapping), synced))

    # ---------- 2 & 3. suffix letters / link text ----------------------------
    fixed_color = 0
    split = 0
    for p in body.findall(".//" + W("w:p")):
        if do_suffix:
            for r in p.findall(W("w:r")):          # direct children only
                if has_link_color(r):
                    rpr = r.find(W("w:rPr"))
                    r.remove(rpr)
                    fixed_color += 1

        if do_split:
            for h in list(p.findall(W("w:hyperlink"))):
                txt = "".join(t.text or "" for t in h.findall(".//" + W("w:t")))
                m = LINKISH.match(txt)
                if not m or not m.group(3):
                    continue
                label, num, tail = m.group(1), m.group(2), m.group(3)
                # keep "Fig. N" inside the link
                for t in h.findall(".//" + W("w:t")):
                    t.text = "%s %s" % (label, num)
                # move the tail out as a plain run right after the hyperlink
                runs = h.findall(W("w:r"))
                if runs:
                    plain = make_plain_run(runs[0], tail)
                    h.addnext(plain)
                    split += 1
    if do_suffix:
        log.append("suffix letters un-styled: %d" % fixed_color)
    if do_split:
        log.append("link texts split: %d" % split)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", help="input .docx (never modified)")
    ap.add_argument("-o", "--out",
                    help="output .docx (default: <name>_fixed.docx next to input)")
    ap.add_argument("--hide-bookmarks", action="store_true")
    ap.add_argument("--fix-suffix-color", action="store_true")
    ap.add_argument("--split-link-text", action="store_true")
    ap.add_argument("--all", action="store_true",
                    help="enable every repair operation")
    ap.add_argument("--dry-run", action="store_true",
                    help="report what would change without writing a file")
    args = ap.parse_args()

    if not any([args.hide_bookmarks, args.fix_suffix_color,
                args.split_link_text, args.all]):
        ap.error("select at least one operation (or --all)")

    doc = docx.Document(args.docx)
    log = []
    fix(doc,
        do_bookmarks=args.all or args.hide_bookmarks,
        do_suffix=args.all or args.fix_suffix_color,
        do_split=args.all or args.split_link_text,
        log=log)

    for line in log:
        print("  %s" % line)

    if args.dry_run:
        print("[dry-run] no file written")
        return 0

    out = args.out
    if not out:
        base, ext = os.path.splitext(args.docx)
        out = base + "_fixed" + ext
    doc.save(out)
    print("[saved] %s" % out)
    print("[next] run format_audit.py on the output to confirm it is clean")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
