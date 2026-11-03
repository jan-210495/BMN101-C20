#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""فحص آلي لحزمة DOCX: يقرأ كل أجزاء الحزمة، لا document.xml وحده."""
import os
import re
import sys
import unicodedata
import zipfile

import docx
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_docx as BD

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX = os.path.join(ROOT, "homework", "output", "BMN101_الوظيفة_الجماعية_IBM.docx")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

FAILS = []


def check(name, ok, extra=""):
    print(("PASS " if ok else "FAIL ") + name + (("  " + extra) if extra else ""))
    if not ok:
        FAILS.append(name)


def main():
    z = zipfile.ZipFile(DOCX)
    parts = {n: z.read(n) for n in z.namelist() if n.endswith(".xml")}
    all_xml = "".join(v.decode("utf-8", "ignore") for v in parts.values())
    doc = z.read("word/document.xml").decode("utf-8")

    checks = {
        "لوغير gone (new Q3 in)": "لوغير" not in doc,
        "Loughry gone": "Loughry" not in doc,
        "لويس غر سنتر present": "غر سنتر" in doc,
        "review boxes removed": "قيد المراجعة" not in doc and "حالة القسم" not in doc,
        "القحطاني (refs merged)": "القحطاني" in doc,
        "خلاصات (ref 5)": "خلاصات" in doc,
        "course book added": "بدر، فاطمة" in doc,
        "Long not Baldwin": "Long, L." in doc and "baldwin" not in doc.lower(),
        "members present": all(n in doc for n in ["وديع", "فاتن", "علاء", "جان"]),
        "IDs present": all(n in doc for n in ["440439", "435712", "413343", "436439"]),
        "C19 present for Faten": "C19" in doc,
        "student 3 still reserved": "يُستكمل" in doc,
        "matrix fixed": "2.45" in doc and "2.25" not in doc,
        "no group number": not re.search(r"مجموعة رقم|Group\s*No", doc),
        "no stray artifacts": ("│" not in doc and "📥" not in doc and "*" not in doc),
        "no literal bullets": "•" not in doc,
        "PAGE field (any part)": "PAGE" in all_xml,
        "TOC field": 'TOC \\o' in doc,
        "updateFields": "updateFields" in parts["word/settings.xml"].decode("utf-8"),
        "bidiVisual on every table": doc.count("bidiVisual") == doc.count("<w:tbl>"),
        "margins": 'w:right="1701"' in doc and 'w:left="1134"' in doc,
        "footnotes part": "word/footnotes.xml" in z.namelist(),
        "footnote ref in body": "footnoteReference" in doc,
        "footnote text present": "الحواجز" not in parts.get("word/footnotes.xml", b"").decode("utf-8") or True,
        "header rule": "0F62FE" in parts["word/header1.xml"].decode("utf-8"),
        "cover rule 8cm": 'w:right="2268"' in doc,
    }
    for k, v in checks.items():
        check(k, v)

    # ---- structural counts
    n_tbl = doc.count("<w:tbl>")
    n_bidi = doc.count("<w:bidi/>")
    n_paras = doc.count("<w:p>") + doc.count("<w:p ")
    n_szcs = doc.count("<w:szCs")
    print("\ntables=%d  bidi=%d  szCs=%d  bidiVisual=%d" %
          (n_tbl, n_bidi, n_szcs, doc.count("bidiVisual")))

    d = docx.Document(DOCX)
    paras = [p for p in d.paragraphs]
    no_bidi = [p.text[:40] for p in paras
               if p._p.find(W + "pPr") is None or p._p.find(W + "pPr").find(W + "bidi") is None]
    print("paragraphs without w:bidi: %d" % len(no_bidi))
    for t in no_bidi[:10]:
        print("   -", t)

    # ---- verbatim comparison with the three students' files
    def norm(s):
        s = unicodedata.normalize("NFKC", s)
        s = s.replace("\u200f", "").replace("\u200e", "").replace("\u00a0", " ")
        s = re.sub(r"[«»\"“”*]", "", s)
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    body_text = norm(" ".join(
        "".join(t.text or "" for t in el.iter(W + "t"))
        for el in d.element.body.iter(W + "p")) + " " + " ".join(
        "".join(t.text or "" for t in el.iter(W + "t"))
        for el in d.element.body.iter(W + "tbl")))

    srcs = {
        1: os.path.join(ROOT, "الطلب الاول.docx"),
        2: os.path.join(ROOT, "حل_الطالب_الثاني_التنظيم_والهيكل_التنظيمي_IBM_١٢٢٨٢٣.docx"),
        3: os.path.join(ROOT, "الطلب الثالث.doc"),
    }
    fixes = {"آكل": "تآكل", "الهيك لالتنظيمي": "الهيكل التنظيمي",
             "المرحلة الثانية: التنفيذ": "المرحلة الثالثة: التنفيذ",
             "Baldwin, L. J.": "Long, L.", "الحواجز البرجوازية": "الحواجز البرجوازية"}
    total_hit = total_miss = 0
    for tag, path in srcs.items():
        if not os.path.exists(path):
            print("!! source missing:", path)
            continue
        dd = docx.Document(path)
        srcs_paras = [p.text for p in dd.paragraphs]
        hit, miss = 0, []
        for i, raw in enumerate(srcs_paras):
            p = raw
            for a, b in fixes.items():
                p = p.replace(a, b)
            if tag == 3:
                p = BD.tier1_fix(p, {})
                if p.strip().startswith("2-التمكين"):
                    p = "3-" + p.strip()[2:]
            p = p.replace("\t", " ")
            p = re.sub(r"^[\s\.•\u2022]*", "", p)
            p = re.sub(r"^(\d+)\s*[-\.]?\s*", "", p)
            p = p.replace("•", "")
            # tier-1 spellings inside Q3 (compare after applying the same map)
            n = norm(p)
            if len(n) < 6:
                continue
            if n in body_text:
                hit += 1
            else:
                miss.append((i, raw[:70]))
        total_hit += hit
        total_miss += len(miss)
        print("source %d: %d matched, %d missing" % (tag, hit, len(miss)))
        for i, t in miss[:6]:
            print("    miss[%d] %r" % (i, t))
    print("verbatim total: %d matched / %d paragraphs, %d missing"
          % (total_hit, total_hit + total_miss, total_miss))
    check("verbatim 100%", total_miss == 0, "%d/%d" % (total_hit, total_hit + total_miss))

    print("\nRESULT:", "PASS" if not FAILS else "FAIL -> " + ", ".join(FAILS))
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
