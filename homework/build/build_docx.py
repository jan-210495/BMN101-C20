#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BMN101 — الوظيفة الجماعية: سكربت الإتمام (سدّ الفجوات السبع).

يعمل على المسوّدة القائمة في homework/output/ ويعدّلها في مكانها:
  * الفجوة 1: استبدال جسم الطلب الثالث بنص الطالب الجديد + تصحيحات الطبقة 1.
  * الفجوة 2: دمج مراجع الطالب الرابع (8 مراجع، قسمان، مسافة معلّقة 1.27 سم).
  * الفجوة 3: بيانات الأعضاء (الاسم/المعرّف/الشعبة) في جدولَي الأعضاء والملحق.
  * الفجوة 4: تصحيح مجموعَي البديلَين (A) و(C) في مصفوفة القرار.
  * الفجوة 5: قوائم الطلب الأول -> ترقيم Word تلقائي.
  * الفجوة 6: مخطط تنظيمي بروابط سلطة رأسية وخط منقّط لفرق المشروعات.
  * الفجوة 7: مسافة معلّقة للمراجع، خط الرأس الأفقي، صناديق «من المصدر».

مفتاح واحد لكل تصحيحات الإملاء في الطلب الثالث:
    SPELLING_TIER1 = "on"   -> تُطبَّق
    SPELLING_TIER1 = "off"  -> تُلغى كلها دفعة واحدة (يبقى النص كما كتبه الطالب)
"""
import os
import re
import sys
import zipfile
from copy import deepcopy
from xml.sax.saxutils import escape

from lxml import etree

SPELLING_TIER1 = "on"

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX = os.path.join(ROOT, "homework", "output", "BMN101_الوظيفة_الجماعية_IBM.docx")
Q3_SRC = os.path.join(ROOT, "الطلب الثالث.doc")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
XML_NS = "http://www.w3.org/XML/1998/namespace"

BLUE = "0F62FE"
NAVY = "001141"
GRAY = "525252"
LIGHT = "E8F0FE"
NEUTRAL = "F4F4F4"
TEXT = "161616"
LINE = "8D8D8D"
WHITE = "FFFFFF"

LOG = []


def log(msg):
    LOG.append(msg)
    print(msg)


# ----------------------------------------------------------------- XML helpers
def q(t):
    return 'w:' + t


def frag(xml_str):
    """Wrap a bare fragment so the w: prefix is bound, and return its root element."""
    if isinstance(xml_str, etree._Element):
        return xml_str
    wrapper = ('<w:root xmlns:w="%s" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               '%s</w:root>' % (W_NS, xml_str))
    return etree.fromstring(wrapper.encode())[0]


def run(text, bold=False, italic=False, color=TEXT, size=32, rtl=True, sup=False,
        style=None):
    rpr = []
    rpr.append('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
               'w:cs="Traditional Arabic" w:eastAsia="Traditional Arabic"/>')
    if style:
        rpr.insert(0, '<w:rStyle w:val="%s"/>' % style)
    rpr.append('<w:b/>' if bold else '<w:b w:val="0"/>')
    if bold:
        rpr.append('<w:bCs/>')
    rpr.append('<w:i/>' if italic else '<w:i w:val="0"/>')
    rpr.append('<w:color w:val="%s"/>' % color)
    rpr.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size))
    if sup:
        rpr.append('<w:vertAlign w:val="superscript"/>')
    if rtl:
        rpr.append('<w:rtl/>')
    body = escape(text)
    return ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
            % ("".join(rpr), body))


def footnote_ref(fid):
    return ('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>'
            '<w:vertAlign w:val="superscript"/></w:rPr>'
            '<w:footnoteReference w:id="%d"/></w:r>' % fid)


def para_xml(runs_xml, style=None, jc="both", bidi=True, before=None, after=120,
             line=360, keep_next=False, keep_lines=False, page_break=False,
             shading=None, numid=None, ind_right=None, ind_left=None,
             hanging=None, outline=None, contextual=False, borders=None,
             tabs=None, rtl_gutter=False):
    p = []
    if style:
        p.append('<w:pStyle w:val="%s"/>' % style)
    if keep_next:
        p.append('<w:keepNext/>')
    if keep_lines:
        p.append('<w:keepLines/>')
    if page_break:
        p.append('<w:pageBreakBefore/>')
    if not rtl_gutter:
        p.append('<w:widowControl/>')
    if numid:
        p.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % numid)
    if borders:
        p.append(borders)
    if shading:
        p.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shading)
    if tabs:
        p.append(tabs)
    if bidi:
        p.append('<w:bidi/>')
    sp = []
    if before is not None:
        sp.append('w:before="%d"' % before)
    if after is not None:
        sp.append('w:after="%d"' % after)
    if line:
        sp.append('w:line="%d" w:lineRule="auto"' % line)
    if sp:
        p.append('<w:spacing %s/>' % " ".join(sp))
    ind = []
    if ind_right is not None:
        ind.append('w:right="%d"' % ind_right)
    if ind_left is not None:
        ind.append('w:left="%d"' % ind_left)
    if hanging is not None:
        ind.append('w:hanging="%d"' % hanging)
    if ind:
        p.append('<w:ind %s/>' % " ".join(ind))
    if contextual:
        p.append('<w:contextualSpacing/>')
    if jc:
        p.append('<w:jc w:val="%s"/>' % jc)
    if outline is not None:
        p.append('<w:outlineLvl w:val="%d"/>' % outline)
    return '<w:p><w:pPr>%s</w:pPr>%s</w:p>' % ("".join(p), runs_xml)


def P(text="", **kw):
    runs = kw.pop("runs", None)
    if runs is None:
        runs = run(text,
                   bold=kw.pop("bold", False),
                   italic=kw.pop("italic", False),
                   color=kw.pop("color", TEXT),
                   size=kw.pop("size", 32),
                   rtl=kw.pop("rtl", True))
    return frag(para_xml(runs, **kw))


def cell(text, width, bold=False, size=28, color=TEXT, fill=None, span=1,
         borders=None, mar=(85, 85, 85, 85), jc="both", valign="center",
         bidi=True, runs=None, line=300, after=40, paras=None):
    tcpr = ['<w:tcW w:type="dxa" w:w="%d"/>' % width]
    if span > 1:
        tcpr.append('<w:gridSpan w:val="%d"/>' % span)
    if borders:
        tcpr.append(borders)
    if fill:
        tcpr.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill)
    tcpr.append('<w:tcMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/>'
                '<w:bottom w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>'
                % (mar[0], mar[1], mar[2], mar[3]))
    if valign:
        tcpr.append('<w:vAlign w:val="%s"/>' % valign)
    if paras is not None:
        body = "".join(paras)
    else:
        if runs is None:
            runs = run(text, bold=bold, color=color, size=size)
        body = para_xml(runs, jc=jc, after=after, line=line, bidi=bidi)
    return '<w:tc><w:tcPr>%s</w:tcPr>%s</w:tc>' % ("".join(tcpr), body)


def tr(cells, height=None, header=False, cant_split=True):
    pr = []
    if cant_split:
        pr.append('<w:cantSplit/>')
    if height:
        pr.append('<w:trHeight w:val="%d" w:hRule="atLeast"/>' % height)
    if header:
        pr.append('<w:tblHeader/>')
    return '<w:tr>%s%s</w:tr>' % ('<w:trPr>%s</w:trPr>' % "".join(pr) if pr else "", "".join(cells))


def table_xml(grid, rows, borders=None, mar=(85, 85, 85, 85), layout="fixed"):
    tblpr = ['<w:tblW w:type="dxa" w:w="%d"/>' % sum(grid), '<w:bidiVisual/>',
             '<w:jc w:val="center"/>']
    if borders:
        tblpr.append(borders)
    if layout:
        tblpr.append('<w:tblLayout w:type="%s"/>' % layout)
    tblpr.append('<w:tblCellMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/>'
                 '<w:bottom w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tblCellMar>'
                 % (mar[0], mar[1], mar[2], mar[3]))
    gridx = "".join('<w:gridCol w:w="%d"/>' % g for g in grid)
    return ('<w:tbl><w:tblPr>%s</w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
            % ("".join(tblpr), gridx, "".join(rows)))


def tbl_borders(outer=(BLUE, 8), inner=(LIGHT, 4)):
    parts = []
    for name in ("top", "left", "bottom", "right"):
        parts.append('<w:%s w:val="single" w:sz="%d" w:space="0" w:color="%s"/>'
                     % (name, outer[1], outer[0]))
    for name in ("insideH", "insideV"):
        parts.append('<w:%s w:val="single" w:sz="%d" w:space="0" w:color="%s"/>'
                     % (name, inner[1], inner[0]))
    return '<w:tblBorders>%s</w:tblBorders>' % "".join(parts)


def borders_all(color=LINE, sz=4, val="single"):
    b = ''.join('<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="%s"/>' % (s, val, sz, color)
                for s in ("top", "left", "bottom", "right"))
    return '<w:tcBorders>%s</w:tcBorders>' % b


def borders_edges(top=None, bottom=None, left=None, right=None):
    parts = []
    for name, spec in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        if spec:
            val, sz, color = spec
            parts.append('<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="%s"/>'
                         % (name, val, sz, color))
        else:
            parts.append('<w:%s w:val="nil"/>' % name)
    return '<w:tcBorders>%s</w:tcBorders>' % "".join(parts)


# ------------------------------------------------------------- package access
class Pkg:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
            self.order = z.namelist()
        self.doc = etree.fromstring(self.parts["word/document.xml"])
        self.body = self.doc.find(W + "body")

    def save(self, path=None):
        self.parts["word/document.xml"] = etree.tostring(
            self.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(path or self.path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.parts[n])
            for n in self.parts:
                if n not in self.order:
                    z.writestr(n, self.parts[n])


def kids(pkg):
    return [k for k in pkg.body if k.tag in (W + "p", W + "tbl")]


def ptext(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def find(pkg, needle, start=0, exact=False):
    ks = kids(pkg)
    for i in range(start, len(ks)):
        t = ptext(ks[i])
        if (t.strip() == needle) if exact else (needle in t):
            return i
    raise KeyError("not found: %r" % needle)


def replace(pkg, first, last, new_xml_list):
    ks = kids(pkg)
    anchor = ks[first]
    for x in new_xml_list:
        anchor.addprevious(frag(x))
    for k in ks[first:last + 1]:
        pkg.body.remove(k)


def insert_after(pkg, idx, new_xml_list):
    ks = kids(pkg)
    anchor = ks[idx]
    for x in reversed(new_xml_list):
        anchor.addnext(frag(x))


# --------------------------------------------------------- schema child order
ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
            "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
            "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
            "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
            "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
            "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
              "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
              "tblPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
             "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "cellIns",
             "cellDel", "cellMerge", "tcPrChange"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
             "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del",
             "trPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
               "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
               "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
               "docGrid", "printerSettings", "sectPrChange"],
}


def normalize_tree(root):
    fixed = 0
    for name, order in ORDER.items():
        idx = {n: i for i, n in enumerate(order)}
        for el in root.iter(W + name):
            children = list(el)
            keyed = [(idx.get(c.tag.split("}")[1], 999), i, c) for i, c in enumerate(children)]
            srt = sorted(keyed, key=lambda x: (x[0], x[1]))
            if [c for _, _, c in srt] != children:
                for c in children:
                    el.remove(c)
                for _, _, c in srt:
                    el.append(c)
                fixed += 1
    return fixed


# ------------------------------------------------------------ Q3 source text
TIER1 = [
    ("ان", "أن", "word", "همزة"),
    ("الى", "إلى", "word", "همزة"),
    ("اجل", "أجل", "word", "همزة"),
    ("او", "أو", "word", "همزة"),
    ("الغاء", "إلغاء", "word", "همزة"),
    ("انشاء", "إنشاء", "word", "همزة"),
    ("انقاذ", "إنقاذ", "word", "همزة"),
    ("ارسال", "إرسال", "word", "همزة"),
    ("اشراك", "إشراك", "word", "همزة"),
    ("واعطائها", "وإعطائها", "word", "همزة"),
    ("اداره", "إدارة", "word", "همزة"),
    ("اعاده", "إعادة", "word", "همزة"),
    ("الازمه", "الأزمة", "word", "همزة"),
    ("ازمه", "أزمة", "word", "همزة"),
    ("الاجمالي", "الإجمالي", "word", "همزة"),
    ("ماليه", "مالية", "word", "تاء مربوطة"),
    ("تنظيميه", "تنظيمية", "word", "تاء مربوطة"),
    ("ثقافه", "ثقافة", "word", "تاء مربوطة"),
    ("خطه", "خطة", "word", "تاء مربوطة"),
    ("طريقه", "طريقة", "word", "تاء مربوطة"),
    ("رؤيه", "رؤية", "word", "تاء مربوطة"),
    ("مباشره", "مباشرة", "word", "تاء مربوطة"),
    ("الكترونيه", "إلكترونية", "word", "تاء مربوطة"),
    ("السريعه", "السريعة", "word", "تاء مربوطة"),
    ("مبتكره", "مبتكرة", "word", "تاء مربوطة"),
    ("مكثفه", "مكثفة", "word", "تاء مربوطة"),
    ("مرنه", "مرنة", "word", "تاء مربوطة"),
    ("مستمره", "مستمرة", "word", "تاء مربوطة"),
    ("العلية", "العليا", "word", "ألف مقصورة"),
    ("حواريه", "حوارية", "word", "تاء مربوطة"),
    ("قادره", "قادرة", "word", "تاء مربوطة"),
    ("تحفيزيه", "تحفيزية", "word", "تاء مربوطة"),
    ("الشركه", "الشركة", "word", "تاء مربوطة"),
    ("العمليه", "العملية", "word", "تاء مربوطة"),
    ("وشاركه", "وشاركة", "word", "تاء مربوطة"),
    ("تكنلوجية", "تكنولوجية", "word", "إملاء"),
    ("البروتكولات", "البروتوكولات", "word", "إملاء"),
    ("التغير", "التغيير", "word", "توحيد المصطلح"),
    ("واداره", "وإدارة", "word", "همزة"),
    ("ازمة", "أزمة", "word", "همزة"),
    ("برؤيه", "برؤية", "word", "تاء مربوطة"),
    ("وطريقه", "وطريقة", "word", "تاء مربوطة"),
    ("وقادره", "وقادرة", "word", "تاء مربوطة"),
    ("والبروتكولات", "والبروتوكولات", "word", "إملاء"),
    ("عمليه", "عملية", "word", "تاء مربوطة"),
]
AR = "\u0621-\u064A\u0640"


def tier1_fix(text, counter):
    if SPELLING_TIER1 != "on":
        return text
    # "الصلاحيات الازمه لاتخاذ" = الصلاحيات اللازمة (لا «الأزمة») -> قاعدة سياقية
    text, n = re.subn(r"الصلاحيات الازمه(?=\s+لاتخاذ)", "الصلاحيات اللازمة", text)
    if n:
        counter[("الازمه (سياق: لاتخاذ)", "اللازمة", "همزة/سياق")] = n
    for src, dst, mode, cat in TIER1:
        pat = r"(?<![%s])%s(?![%s])" % (AR, re.escape(src), AR)
        text, n = re.subn(pat, dst, text)
        if n:
            counter[(src, dst, cat)] = counter.get((src, dst, cat), 0) + n
    return text


def load_q3():
    import docx
    d = docx.Document(Q3_SRC)
    out = []
    for p in d.paragraphs:
        t = p.text.replace("\t", " ").strip()
        if t:
            out.append((p.style.name, t))
    return out


# ----------------------------------------------------------------- the edits
def main():
    pkg = Pkg(DOCX)

    # ============================ GAP 4 — decision matrix totals
    mat = None
    for t in pkg.body.iter(W + "tbl"):
        if "المجموع الموزون" in ptext(t):
            mat = t
            break
    if mat is None:
        raise SystemExit("matrix not found")
    rows = mat.findall(W + "tr")
    last = rows[-1]
    tcs = last.findall(W + "tc")
    vals = []
    for tc in tcs:
        ts = [x.text for x in tc.iter(W + "t")]
        vals.append("".join(ts))
    log("matrix before: %s" % vals)
    # columns: المعيار | الوزن | A | B | C | D
    assert vals[2].strip() == "2.25", vals
    assert vals[4].strip() == "2.35", vals
    tcs[2].iter(W + "t").__next__().text = "2.35"
    tcs[4].iter(W + "t").__next__().text = "2.45"
    # keep run props consistent (first run already plain text)
    log("matrix after : A=2.35 C=2.45")

    # ============================ GAP 3 — members tables
    members = [
        ("1", "وديع هيثم سلوم", "wadee_440439", "C20",
         "الطلب الأول: التخطيط واتخاذ القرار", "تحليل المشكلة والبدائل والمعايير والقرار"),
        ("2", "فاتن سلمان عبد الولي", "Faten_435712", "C19",
         "الطلب الثاني: التنظيم والهيكل التنظيمي", "تحليل المشكلات التنظيمية والمخطط المقترح"),
        ("3", "⟨يُستكمل⟩", "⟨يُستكمل⟩", "⟨يُستكمل⟩",
         "الطلب الثالث: القيادة والتوجيه والتحفيز وإدارة التغيير", "خطة إدارة التغيير"),
        ("4", "علاء اسامه الحايك", "alaa_413343", "C20",
         "المراجع والتوثيق العلمي", "قائمة المراجع وتدقيق الاستشهادات"),
        ("5", "جان شادي عبود", "jan_436439", "C20",
         "التنسيق والإخراج والإبداع والدمج النهائي", "الملف النهائي بصيغتَيه"),
    ]
    GRID = [430, 1750, 1180, 620, 2740, 2351]  # 9071

    def members_table():
        out = []
        head = ["#", "اسم العضو", "المعرّف", "الشعبة", "المهمة", "المخرج"]
        out.append(tr([cell(h, GRID[i], bold=True, color=WHITE, fill=BLUE, size=28, jc="center")
                       for i, h in enumerate(head)], height=420, header=True))
        for k, row in enumerate(members):
            fill = NEUTRAL if k % 2 else None
            cells = []
            for i, v in enumerate(row):
                ph = v == "⟨يُستكمل⟩"
                cells.append(cell(v, GRID[i], size=26,
                                  fill=("F4F4F4" if ph else fill),
                                  bold=(i == 1),
                                  jc="center" if i in (0, 2, 3) else "both"))
            out.append(tr(cells))
        return table_xml(GRID, out, borders=tbl_borders())

    n_mem = 0
    for t in list(pkg.body.iter(W + "tbl")):
        hdr = [ptext(tc).strip() for tc in t.findall(W + "tr")[0].findall(W + "tc")]
        if hdr[:3] == ["#", "اسم العضو", "الرقم الجامعي"]:
            new = frag(members_table())
            t.addprevious(new)
            pkg.body.remove(t)
            n_mem += 1
    log("members tables rebuilt: %d" % n_mem)

    # cover: no single section label, logo placeholder kept
    cover_txt = ptext(kids(pkg)[0]) + ptext(kids(pkg)[10])
    assert "C20" not in cover_txt and "C19" not in cover_txt

    # cover rule 8 cm (4535 twips), centered
    side = (9071 - 4535) // 2
    rule = para_xml("", after=240, before=0, borders=(
        '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="%s"/></w:pBdr>' % BLUE),
        ind_right=side, ind_left=side, jc="center", line=240)
    i_title = find(pkg, "التحول الإداري في شركة IBM", 0, exact=True)
    insert_after(pkg, i_title, [rule])
    log("cover rule added (8 cm)")

    # ============================ GAP 1 — Q3 replacement
    q3 = load_q3()
    counter = {}
    fixed_q3 = [(st, tier1_fix(tx, counter)) for st, tx in q3]
    log("Q3 paragraphs: %d (source non-empty)" % len(fixed_q3))
    log("tier-1 source chars: %d -> fixed chars: %d"
        % (sum(len(t) for _, t in q3), sum(len(t) for _, t in fixed_q3)))

    # duplicate numbering in ثانياً
    for i, (st, tx) in enumerate(fixed_q3):
        if st == "Heading 3" and tx.strip().startswith("2-التمكين"):
            fixed_q3[i] = (st, "3-" + tx.strip()[2:])
            log("duplicate item number fixed: 2-التمكين -> 3-التمكين")

    # numbering.xml: RTL decimal list, three instances for three restarts
    numbering = etree.fromstring(pkg.parts["word/numbering.xml"])
    abs_id = 20
    lvl = (
        '<w:abstractNum w:abstractNumId="%d">'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%%1."/><w:lvlJc w:val="right"/>'
        '<w:pPr><w:ind w:right="425" w:hanging="425"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Traditional Arabic"/><w:b w:val="0"/><w:i w:val="0"/>'
        '<w:sz w:val="32"/><w:szCs w:val="32"/><w:rtl/></w:rPr>'
        '</w:lvl></w:abstractNum>' % abs_id)
    first_num = numbering.find(W + "num")
    abs_el = frag(lvl)
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        numbering.append(abs_el)
    for nid in (31, 32, 33):
        numbering.append(frag(
            '<w:num w:numId="%d"><w:abstractNumId w:val="%d"/>'
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
            '</w:num>' % (nid, abs_id)))
    pkg.parts["word/numbering.xml"] = etree.tostring(numbering, xml_declaration=True,
                                                    encoding="UTF-8", standalone=True)

    # build the new Q3 paragraphs
    new_q3 = []
    h2n = h3n = 0
    body_no = 0
    fn_added = False
    for st, tx in fixed_q3:
        if st == "Heading 1":
            new_q3.append(P(tx, size=36, bold=True, color=NAVY, jc="right", before=0, after=180,
                            keep_next=True, line=330))
            continue
        if st == "Heading 2":
            h2n += 1
            new_q3.append(P(tx, style="Heading2", size=34, bold=True, color=NAVY, jc="right",
                            before=240, after=120, keep_next=True, line=330, outline=1))
            continue
        if st == "Heading 3":
            h3n += 1
            new_q3.append(P(tx, style="Heading3", size=32, bold=True, color=BLUE, jc="right",
                            before=180, after=80, keep_next=True, line=330, outline=2))
            continue
        body_no += 1
        if "الحواجز البرجوازية" in tx:
            head, tail = tx.split("الحواجز البرجوازية", 1)
            runs = (run(head + "الحواجز البرجوازية", size=32) + footnote_ref(1)
                    + run(tail, size=32))
            new_q3.append(P(runs=runs, after=120))
            fn_added = True
        else:
            new_q3.append(P(tx, after=120))
    log("Q3 built: %d H2, %d H3, %d body paragraphs" % (h2n, h3n, body_no))
    if fn_added:
        add_footnote(pkg)

    first = find(pkg, "⚠ حالة القسم")
    last = find(pkg, "شركة نتفليكس")
    log("Q3 span: %d..%d (%d elements replaced)" % (first, last, last - first + 1))
    replace(pkg, first, last, new_q3)
    assert "لوغير" not in etree.tostring(pkg.doc, encoding="unicode")

    # ============================ GAP 5 + literal bullets in Q1
    def strip_lead(t):
        t2 = re.sub(r"^[\s\u200f\.\u00a0]*", "", t)
        t2 = re.sub(r"^\d+\s*[\.\u2013-]?\s*", "", t2)
        t2 = re.sub(r"^[\s•\u2022]+", "", t2)
        return t2.strip()

    def listify(pkg, first_i, last_i, nid):
        ks = kids(pkg)
        for k in ks[first_i:last_i + 1]:
            t = ptext(k)
            new = strip_lead(t)
            assert new, repr(t)
            for tt in k.iter(W + "t"):
                tt.text = ""
            runs = k.findall(W + "r")
            target = runs[-1] if runs else None
            if target is None:
                continue
            for r in runs[:-1]:
                k.remove(r)
            tels = target.findall(W + "t")
            if not tels:
                tels = [etree.SubElement(target, W + "t")]
            for t_el in tels:
                t_el.text = new
                t_el.set("{%s}space" % XML_NS, "preserve")
            # numbering properties
            ppr = k.find(W + "pPr")
            if ppr is None:
                ppr = etree.SubElement(k, W + "pPr")
                k.remove(ppr)
                k.insert(0, ppr)
            for old in ppr.findall(W + "numPr"):
                ppr.remove(old)
            numpr = frag('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % nid)
            ppr.insert(0, numpr)
            jc = ppr.find(W + "jc")
            if jc is not None:
                jc.set(W + "val", "both")

    for needle_start, needle_end, nid, label in (
            ("خسائر مالية ضخمة", "إذن المشكلة الجوهرية", 31, "problem list"),
            ("القدرة على إعادة بناء القدرة", "تأثير القرار على الثقافة", 32, "criteria list"),
            ("الأزمة مرتبطة بالتنسيق", "تغيير الثقافة والإدارة ضروري", 33, "justifications"),
    ):
        a = find(pkg, needle_start)
        b = find(pkg, needle_end, a)
        listify(pkg, a, b, nid)
        log("numbered list (%s): %d items" % (label, b - a + 1))

    # literal bullets -> List Bullet paragraphs
    nb = 0
    for k in list(pkg.body.iter(W + "p")):
        if k.tag != W + "p":
            continue
        t = ptext(k)
        if "•" not in t:
            continue
        txt = t.replace("•", "").strip()
        if not txt:
            continue
        for tt in k.iter(W + "t"):
            tt.text = ""
        runs = k.findall(W + "r")
        if not runs:
            continue
        for r in runs[:-1]:
            k.remove(r)
        for t_el in runs[-1].findall(W + "t"):
            t_el.text = txt
            t_el.set("{%s}space" % XML_NS, "preserve")
        ppr = k.find(W + "pPr")
        if ppr.find(W + "pStyle") is None:
            ppr.insert(0, frag('<w:pStyle w:val="ListBullet"/>'))
        nb += 1
    log("literal bullets converted to List Bullet: %d" % nb)


    # ---------------------------- Q1/Q2 subsection headings -> Heading 2
    def as_h2(el):
        ppr = el.find(W + "pPr")
        if ppr is None:
            return
        for old in ppr.findall(W + "pStyle"):
            ppr.remove(old)
        ppr.insert(0, frag('<w:pStyle w:val="Heading2"/>'))
        for name, val in (("keepNext", None),):
            if ppr.find(W + name) is None:
                ppr.append(frag('<w:%s/>' % name))
        sp = ppr.find(W + "spacing")
        if sp is not None:
            sp.set(W + "before", "240")
            sp.set(W + "after", "120")
        jc = ppr.find(W + "jc")
        if jc is not None:
            jc.set(W + "val", "right")
        for r in el.findall(W + "r"):
            rpr = r.find(W + "rPr")
            if rpr is None:
                continue
            b = rpr.find(W + "b")
            if b is None:
                rpr.append(frag('<w:b/>'))
            else:
                b.set(W + "val", "1")
            if rpr.find(W + "bCs") is None:
                rpr.append(frag('<w:bCs/>'))
            c = rpr.find(W + "color")
            if c is not None:
                c.set(W + "val", NAVY)
            sz = rpr.find(W + "sz")
            if sz is not None:
                sz.set(W + "val", "34")
            szcs = rpr.find(W + "szCs")
            if szcs is not None:
                szcs.set(W + "val", "34")

    Q1_HEADS = ("المشكلة (Problem Statement)", "البدائل المتاحة (Alternatives)",
                "معايير اتخاذ القرار (Decision Criteria)", "تقييم موجز للبدائل وفق المعايير",
                "القرار المقترح (Recommended Decision)", "مبررات واضحة (Justifications)",
                "خاتمة قصيرة")
    Q2_HEADS = ("تحليل المشكلات التنظيمية في الحالة", "الهيكل التنظيمي المقترح",
                "مهام الوحدات الرئيسية ومسؤولياتها", "العلاقات التنظيمية وخطوط السلطة والتنسيق",
                "مبررات اختيار الهيكل المقترح", "النتائج المتوقعة من تطبيق الهيكل", "الخلاصة")
    n_h2 = 0
    for el in pkg.body.iter(W + "p"):
        t = ptext(el).strip()
        core = re.sub(r"^[\s\d\.]*", "", t)
        if any(core.startswith(h) for h in Q1_HEADS) or t in Q2_HEADS:
            as_h2(el)
            n_h2 += 1
    log("subsection headings promoted to Heading 2: %d" % n_h2)

    # ---------------------------- broken word inside the student's Q2 text
    # «الهيك لالتنظيمي» كلمة واحدة كُسرت بمسافة داخلية تمنع ظهورها في Word
    n_broken = 0
    for t in pkg.body.iter(W + "t"):
        if t.text and "الهيك لالتنظيمي" in t.text:
            t.text = t.text.replace("الهيك لالتنظيمي", "الهيكل التنظيمي")
            n_broken += 1
    if n_broken:
        log("broken word repaired: «الهيك لالتنظيمي» -> «الهيكل التنظيمي» (×%d)" % n_broken)

    # ============================ GAP 2 — references
    refs_ar = [
        ("بدر، فاطمة. (2018). ", "أساسيات الإدارة", ". منشورات الجامعة الافتراضية السورية."),
        ("جرستنر، لويس. (2003، ديسمبر). من قال إن الفيلة لا تستطيع الرقص؟ كيف نجحت إدارة IBM "
         "في إنقاذ شركتها من السقوط [ملخّص]. ", "خلاصات كتب المدير ورجل الأعمال",
         "، الشركة العربية للإعلام العلمي (شعاع)، 1–8."),
        ("القحطاني، سالم بن سعيد. (2008). ", "القيادة الإدارية: التحوّل نحو نموذج القيادة العالمي",
         " (ط2). مرامر للطباعة والتغليف."),
    ]
    refs_en = [
        ("Gerstner, L. V., Jr. (2002). ", "Who says elephants can't dance? Inside IBM's historic turnaround", ". HarperBusiness."),
        ("IBM. (n.d.). ", "IBM history: Louis V. Gerstner Jr.", " Retrieved November 3, 2026, from https://www.ibm.com/history/louis-gerstner"),
        ("IBM. (n.d.). ", "IBM's business transformation unfolds", ". Retrieved November 3, 2026, from http://www.public.dhe.ibm.com/software/solutions/pdfs/g325-1931-00.pdf"),
        ("Kotter, J. P. (1996). ", "Leading change", ". Harvard Business School Press."),
        ("Robbins, S. P., Coulter, M. A., & Long, L. (2024). ", "Management", " (16th ed.). Pearson."),
    ]

    def ar_ref(runs):
        return para_xml(runs, jc="both", after=120, ind_right=720, hanging=720, line=360)

    def en_ref(runs):
        return ('<w:p><w:pPr><w:widowControl/><w:spacing w:after="120" w:line="360" '
                'w:lineRule="auto"/><w:ind w:left="720" w:hanging="720"/><w:jc w:val="left"/>'
                '</w:pPr>%s</w:p>' % runs)

    # remove reserved box on the references page
    i_refbox = None
    for i, k in enumerate(kids(pkg)):
        if k.tag == W + "tbl" and "مكان محجوز" in ptext(k) and "المراجع" in ptext(k):
            i_refbox = i
    if i_refbox is None:
        for i, k in enumerate(kids(pkg)):
            if k.tag == W + "tbl" and "مكان محجوز — إعداد الطالب الرابع" in ptext(k):
                i_refbox = i
                break
    ks = kids(pkg)
    nxt_idx = i_refbox + 1
    while ks[nxt_idx].tag != W + "p" or not ptext(ks[nxt_idx]).strip().startswith("المراجع العربية"):
        nxt_idx += 1
    pkg.body.remove(ks[i_refbox])
    log("references reserved box removed")

    # Arabic section body
    a = find(pkg, "بدر، فاطمة. (2018)")
    b = find(pkg, "المراجع الأجنبية", a) - 1
    ar_xml = [ar_ref(run(p1, size=30) + run(mid, size=30, italic=True) + run(p2, size=30))
              for p1, mid, p2 in refs_ar]
    replace(pkg, a, b, ar_xml)

    # Foreign section body
    a = find(pkg, "المراجع الأجنبية") + 1
    b = find(pkg, "ملحق (أ)", a) - 1
    en_xml = [en_ref(run(p1, size=30, rtl=False) + run(mid, size=30, italic=True, rtl=False)
                     + run(p2, size=30, rtl=False)) for p1, mid, p2 in refs_en]
    replace(pkg, a, b, en_xml)
    log("references rebuilt: 3 Arabic + 5 foreign")

    # ============================ GAP 7 — «من المصدر» boxes
    def source_box(lines):
        inner = []
        for k, line_ in enumerate(lines):
            inner.append(para_xml(
                run(line_, bold=(k == 0), color=(GRAY if k == 0 else TEXT), size=(30 if k == 0 else 28)),
                jc="both", after=(60 if k == 0 else 40), line=300))
        return table_xml([9071], [tr([cell("", 9071, fill=NEUTRAL,
                                           borders=borders_all(LINE, 4),
                                           paras=inner, mar=(100, 140, 100, 140))])],
                         borders=None)

    box_q1 = source_box([
        "من المصدر — الطلب الأول",
        "الوقائع والأرقام (خسارة 8.10 مليار دولار عام 1993، وإيرادات 62.7 مليار دولار) مأخوذة من ورقة الوظيفة "
        "ومن Gerstner (2002)؛ ومفاهيم التخطيط واتخاذ القرار والمعايير الموزونة من كتاب المقرر بدر (2018، الفصل "
        "السابع) ومن Robbins, Coulter & Long (2024). أما الأوزان والدرجات في المصفوفة فهي تقدير المجموعة، "
        "وقد أُفصح عن ذلك في صندوق التحليل.",
    ])
    box_q2 = source_box([
        "من المصدر — الطلب الثاني",
        "المفاهيم التنظيمية المستخدمة (التخصص، نطاق الإشراف، التفويض، اللامركزية، التنسيق) من بدر (2018، الفصل "
        "الثامن) ومن Robbins, Coulter & Long (2024)؛ ووقائع تشتّت وحدات IBM ومشكلات التنسيق بينها من "
        "Gerstner (2002) ومن IBM (بلا تاريخ). والهيكل المرسوم في الشكل (2) اقتراح تحليلي للمجموعة لا نقلٌ عن "
        "هيكل IBM التاريخي.",
    ])
    box_q3 = source_box([
        "من المصدر — الطلب الثالث",
        "مراحل التغيير الثلاث الواردة في النص (خلق شعور بالجدية والاضطرار، بناء تحالف موجّه، تسييل العادة "
        "القديمة) تقابل الخطوات الأولى في نموذج Kotter (1996)، وأساليب التعامل مع المقاومة (التعليم والاتصال، "
        "المشاركة والاندماج، التسهيل والدعم التدريبي) تقابل أدبيات إدارة التغيير عند Kotter (1996) والقحطاني "
        "(2008) وبدر (2018)؛ ووقائع تحوّل IBM وقيادة لويس غر سنتر من Gerstner (2002) ومن IBM (بلا تاريخ). "
        "ولم تُحقن استشهادات داخل نصوص الطلاب، بل جُمعت هنا.",
    ])

    # Q1: before the group-analysis box that follows the matrix
    i_grp1 = None
    ks = kids(pkg)
    mat_i = [i for i, k in enumerate(ks) if k.tag == W + "tbl" and "المجموع الموزون" in ptext(k)][0]
    for i in range(mat_i, len(ks)):
        if ks[i].tag == W + "tbl" and "تحليل المجموعة" in ptext(ks[i]):
            i_grp1 = i
            break
    insert_after(pkg, i_grp1 - 1, [box_q1])
    # Q2: before its group box
    ks = kids(pkg)
    h1q3 = [i for i, k in enumerate(ks) if ptext(k).startswith("الطلب الثالث")][0]
    i_grp2 = None
    for i in range(h1q3 - 1, 0, -1):
        if ks[i].tag == W + "tbl" and "تحليل المجموعة" in ptext(ks[i]):
            i_grp2 = i
            break
    insert_after(pkg, i_grp2 - 1, [box_q2])
    # Q3: right after the new student text (before the roadmap table)
    i_fig3 = find(pkg, "المرحلة الأولى: التحضير")
    ks = kids(pkg)
    while ks[i_fig3].tag != W + "tbl":
        i_fig3 -= 1
    insert_after(pkg, i_fig3 - 1, [box_q3])
    log("«من المصدر» boxes inserted: 3")

    # ============================ GAP 6 — org chart
    depts = [
        ("إدارة البرمجيات والحلول الرقمية", "تطوير البرمجيات والمنصات والحلول الرقمية"),
        ("إدارة البنية التحتية والأنظمة التقنية", "توفير الأنظمة التقنية والبنية التحتية وتشغيلها"),
        ("إدارة الخدمات والاستشارات التقنية", "تحليل احتياجات العملاء وتصميم الحلول وتنفيذها"),
        ("إدارة البحث والتطوير والابتكار", "البحث في التقنيات الجديدة وتحويلها إلى حلول قابلة للتطبيق"),
        ("إدارة التسويق والمبيعات وخدمة العملاء", "فهم السوق وإدارة العلاقات وقياس رضا العملاء"),
        ("الإدارات الداعمة المشتركة", "الموارد البشرية والمالية والشؤون القانونية والعمليات"),
    ]
    sub = [721, 70, 721]
    grid = sub * 6
    grid[-1] -= 1
    DBLUE = ("single", 8, BLUE)
    rows = []
    rows.append(tr([cell("الرئيس التنفيذي (CEO)", 9071, bold=True, color=WHITE, fill=BLUE,
                         size=30, jc="center", borders=borders_all(BLUE, 8), span=18)],
                   height=460))
    # bus
    rows.append(tr([cell("", 9071, span=18, size=4,
                         borders=borders_edges(top=None, left=None, right=None, bottom=DBLUE),
                         mar=(0, 0, 0, 0))], height=60))
    # stems: one 70-twip cell per department, left border only
    stem_cells = []
    for _ in range(6):
        stem_cells.append(cell("", 721, size=4, borders=borders_edges(), mar=(0, 0, 0, 0)))
        stem_cells.append(cell("", 70, size=4, borders=borders_edges(left=DBLUE), mar=(0, 0, 0, 0)))
        stem_cells.append(cell("", 721, size=4, borders=borders_edges(), mar=(0, 0, 0, 0)))
    rows.append(tr(stem_cells, height=430))
    # departments
    dcells = []
    for name, desc in depts:
        inner = (para_xml(run(name, bold=True, color=NAVY, size=26), jc="center", after=40, line=280)
                 + para_xml(run(desc, size=24, color=GRAY), jc="center", after=0, line=260))
        dcells.append(cell("", 1512, span=3, fill=LIGHT, borders=borders_all(BLUE, 8),
                           runs=inner, mar=(60, 60, 60, 60), valign="top"))
    rows.append(tr(dcells, height=900))
    # dashed separator
    rows.append(tr([cell("", 9071, span=18, size=4, mar=(0, 0, 0, 0),
                         borders=borders_edges(top=("dashed", 8, GRAY)))], height=120))
    # horizontal coordination band
    rows.append(tr([cell("فرق المشروعات المشتركة (تنسيق أفقي حول العميل)",
                         9071, bold=True, size=26, color=NAVY, fill=NEUTRAL, jc="center",
                         span=18, borders=borders_edges(top=("dashed", 8, GRAY),
                                                        bottom=("dashed", 8, GRAY),
                                                        left=("dashed", 8, GRAY),
                                                        right=("dashed", 8, GRAY)))]))
    org = table_xml(grid, rows, borders=None)
    i_org = find(pkg, "الرئيس التنفيذي (CEO)")
    ks = kids(pkg)
    while ks[i_org].tag != W + "tbl":
        i_org -= 1
    old = ks[i_org]
    old.addprevious(frag(org))
    pkg.body.remove(old)
    log("org chart rebuilt: 18-column grid, stems + dashed coordination band")

    # ============================ annex: member table + coordination sentence
    n_ann = len([1 for t in pkg.body.iter(W + "tbl")
                 if [ptext(tc).strip() for tc in t.findall(W + "tr")[0].findall(W + "tc")][:3]
                 == ["#", "اسم العضو", "المعرّف"]])
    log("member tables present after rebuild (members + annex): %d" % n_ann)

    # ============================ header rule (0.5 pt, #0F62FE)
    hdr = etree.fromstring(pkg.parts["word/header1.xml"])
    hp = hdr.find(W + "p")
    hp_pr = hp.find(W + "pPr")
    for old in hp_pr.findall(W + "pBdr"):
        hp_pr.remove(old)
    pbdr = etree.fromstring(
        ('<w:pBdr xmlns:w="%s"><w:bottom w:val="single" w:sz="4" w:space="1" '
         'w:color="%s"/></w:pBdr>' % (W_NS, BLUE)).encode())
    hp_pr.append(pbdr)
    pkg.parts["word/header1.xml"] = etree.tostring(hdr, xml_declaration=True,
                                                   encoding="UTF-8", standalone=True)
    normalize_tree(hdr)
    log("header bottom rule added (#0F62FE, 0.5 pt)")

    # annex: the sentence about unifying proper names no longer holds (two spellings
    # coexist by design: جيرستنر in the coordinator text, غر سنتر in the student text)
    for el in pkg.body.iter(W + "p"):
        t = ptext(el)
        if "اعتماد صيغة «لو جيرستنر» في المستند كله" in t:
            t2 = t.replace(
                "وُحّدت المصطلحات والأعلام (ومنها اعتماد صيغة «لو جيرستنر» في المستند كله)",
                "وُحّدت المصطلحات والأعلام حيث أمكن، مع الحفاظ على صيغة كل طالب داخل نصّه "
                "(«جيرستنر» في أقسام المنسّق، و«لويس غر سنتر» في نص الطلب الثالث)")
            ts = list(el.iter(W + "t"))
            ts[0].text = t2
            for extra in ts[1:]:
                extra.text = ""
            log("annex name-unification sentence corrected")
            break

    # ============================ GAP 1.4 — Figure 3 (roadmap) rebuilt
    f3_head = ["", "المرحلة الأولى: خلق شعور بالجدية والاضطرار",
               "المرحلة الثانية: بناء تحالف موجّه",
               "المرحلة الثالثة: تسييل وتفكيك العادة القديمة"]
    f3_rows = [
        ("ما يقوله النص",
         "مصارحة الجميع بالخسائر المالية الكبيرة لإدراك خطورة البقاء على الوضع الراهن",
         "تشكيل فريق قيادي مرن يؤمن بالرؤية الجديدة وقادر على نشرها في كافة المستويات التنظيمية",
         "التخلص من القواعد الصارمة والبروتوكولات المعقدة التي كانت تسبب بطء وتيرة العمل اليومي"),
        ("مقابلها في نموذج كوتر (1996)",
         "الخطوة 1: إشعار بالإلحاح",
         "الخطوة 2: بناء تحالف موجّه",
         "الخطوة 3: إزالة العوائق أمام الرؤية الجديدة"),
        ("أسلوب التعامل مع المقاومة المصاحب",
         "التعليم والاتصال المستمر",
         "المشاركة والاندماج",
         "التسهيل والدعم التدريبي"),
        ("مؤشر الانتقال",
         "إدراك مشترك لحجم الأزمة",
         "فريق قيادي يملك الرؤية وينشرها",
         "إجراءات أبسط وقرارات أسرع"),
    ]
    G3 = [1500, 2524, 2524, 2523]
    rows3 = [tr([cell(f3_head[i], G3[i], bold=True, color=WHITE, fill=BLUE, size=26,
                      jc="center") for i in range(4)], height=520, header=True)]
    for k, (label_, *vals) in enumerate(f3_rows):
        fill = NEUTRAL if k % 2 == 0 else None
        cells = [cell(label_, G3[0], bold=True, color=NAVY, fill=LIGHT, size=24, jc="right")]
        for i, v in enumerate(vals, start=1):
            cells.append(cell(v, G3[i], size=24, fill=fill, jc="both"))
        rows3.append(tr(cells))
    fig3 = frag(table_xml(G3, rows3, borders=tbl_borders()))
    for t in list(pkg.body.iter(W + "tbl")):
        if "المرحلة الأولى: التحضير" in ptext(t):
            t.addprevious(fig3)
            pkg.body.remove(t)
            log("Figure 3 rebuilt: 3 stages of the new text + Kotter mapping")
            break

    # ---------------------------- Table 3: align treatments with the new text
    treat = {
        "المشاركة والتضمين في تصميم الهيكل": "المشاركة والاندماج في تصميم الهيكل",
        "الاتصال الشفاف + الدعم والتدريب": "التعليم والاتصال المستمر + الدعم التدريبي",
        "التفاوض + مؤشرات أداء واضحة": "التسهيل والدعم التدريبي + مؤشرات أداء واضحة",
    }
    for t in pkg.body.iter(W + "tbl"):
        if "فئة المقاومة" not in ptext(t):
            continue
        for r in t.findall(W + "tr")[1:]:
            tcs = r.findall(W + "tc")
            if len(tcs) < 3:
                continue
            cur = ptext(tcs[2]).strip()
            if cur in treat:
                ts = list(tcs[2].iter(W + "t"))
                ts[0].text = treat[cur]
                for extra in ts[1:]:
                    extra.text = ""
                log("Table 3 treatment updated: %r -> %r" % (cur, treat[cur]))
        break

    # ---------------------------- TOC styles (RTL, dot leader)
    styles = etree.fromstring(pkg.parts["word/styles.xml"])
    have = {x.get(W + "styleId") for x in styles.findall(W + "style")}
    toc_style = (
        '<w:style xmlns:w="%s" w:type="paragraph" w:styleId="TOC%d">'
        '<w:name w:val="toc %d"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
        '<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
        '<w:pPr><w:tabs><w:tab w:val="left" w:leader="dot" w:pos="8800"/></w:tabs>'
        '<w:bidi/><w:spacing w:after="60" w:line="300" w:lineRule="auto"/>'
        '<w:ind w:right="%d"/><w:jc w:val="right"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Traditional Arabic"/>%s<w:color w:val="%s"/>'
        '<w:sz w:val="%d"/><w:szCs w:val="%d"/><w:rtl/></w:rPr></w:style>')
    for lvl, bold, color, sz, ind in ((1, "<w:b/><w:bCs/>", NAVY, 28, 0),
                                      (2, "", GRAY, 26, 240)):
        if "TOC%d" % lvl not in have:
            styles.append(etree.fromstring(
                (toc_style % (W_NS, lvl, lvl, ind, bold, color, sz, sz)).encode()))
    pkg.parts["word/styles.xml"] = etree.tostring(styles, xml_declaration=True,
                                                 encoding="UTF-8", standalone=True)
    log("TOC1/TOC2 styles added")

    # ---------------------------- annex table: citation row + Q1 numbering artefacts
    for el in pkg.body.iter(W + "p"):
        t = ptext(el)
        if t.strip().startswith("داخل النص بأسلوب APA"):
            ts = list(el.iter(W + "t"))
            ts[0].text = ("صناديق «من المصدر» في كل طلب (لا استشهادات داخل نصوص الطلاب) · "
                          "قائمة مراجع APA 7 بمسافة بادئة معلّقة 1.27 سم")
            for extra in ts[1:]:
                extra.text = ""
            log("annex citation row updated (boxes instead of in-text citations)")

    for el in pkg.body.iter(W + "p"):
        t = ptext(el)
        if t.strip().startswith("1المشكلة"):
            ts = list(el.iter(W + "t"))
            ts[0].text = t.replace("1المشكلة", "1. المشكلة", 1)
            for extra in ts[1:]:
                extra.text = ""
            log("Q1 heading separator restored: «1المشكلة» -> «1. المشكلة»")
        elif t.strip().startswith("6. خاتمة قصيرة"):
            ts = list(el.iter(W + "t"))
            ts[0].text = t.replace("6. خاتمة قصيرة", "7. خاتمة قصيرة", 1).strip()
            for extra in ts[1:]:
                extra.text = ""
            log("Q1 duplicate item number fixed: 6. خاتمة -> 7. خاتمة")

    # ============================ normalize child order + save
    fixed = normalize_tree(pkg.doc)
    log("schema-order fixes: %d elements" % fixed)
    pkg.save()
    log("saved: %s" % DOCX)

    with open("/tmp/spelling-log.txt", "w", encoding="utf-8") as f:
        for (src, dst, cat), n in sorted(counter.items(), key=lambda x: -x[1]):
            f.write("%-14s -> %-14s  ×%-3d  (%s)\n" % (src, dst, n, cat))
        f.write("\nTOTAL: %d\n" % sum(counter.values()))
    log("spelling cases: %d" % sum(counter.values()))


def add_footnote(pkg):
    fn = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:footnotes xmlns:w="%s">'
          '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" '
          'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
          '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing '
          'w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/>'
          '</w:r></w:p></w:footnote>'
          '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/><w:bidi/>'
          '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/>'
          '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
          'w:cs="Traditional Arabic"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr>'
          '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
          '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
          'w:cs="Traditional Arabic"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl/></w:rPr>'
          '<w:t xml:space="preserve"> كذا في الأصل، والمرجّح أن المقصود «البيروقراطية»؛ بدليل '
          'ورود «تم التخلي عن البيروقراطية» في الفقرة التالية.</w:t></w:r>'
          '</w:p></w:footnote></w:footnotes>' % W_NS)
    pkg.parts["word/footnotes.xml"] = fn.encode("utf-8")

    # content types
    ct = etree.fromstring(pkg.parts["[Content_Types].xml"])
    ns = ct.nsmap[None]
    have = [o.get("PartName") for o in ct if o.tag.endswith("Override")]
    if "/word/footnotes.xml" not in have:
        o = etree.SubElement(ct, "{%s}Override" % ns)
        o.set("PartName", "/word/footnotes.xml")
        o.set("ContentType",
              "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")
    pkg.parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True,
                                                     encoding="UTF-8", standalone=True)
    # relationship
    rels = etree.fromstring(pkg.parts["word/_rels/document.xml.rels"])
    rns = rels.nsmap[None]
    ids = {r.get("Id") for r in rels}
    rid = "rIdFootnotes1"
    if rid not in ids:
        r = etree.SubElement(rels, "{%s}Relationship" % rns)
        r.set("Id", rid)
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes")
        r.set("Target", "footnotes.xml")
    pkg.parts["word/_rels/document.xml.rels"] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # footnote styles
    styles = etree.fromstring(pkg.parts["word/styles.xml"])
    ids = {s.get(W + "styleId") for s in styles.findall(W + "style")}
    if "FootnoteText" not in ids:
        styles.append(etree.fromstring((
            '<w:style xmlns:w="%s" w:type="paragraph" w:styleId="FootnoteText">'
            '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>' % W_NS).encode()))
    if "FootnoteReference" not in ids:
        styles.append(etree.fromstring((
            '<w:style xmlns:w="%s" w:type="character" w:styleId="FootnoteReference">'
            '<w:name w:val="footnote reference"/>'
            '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' % W_NS).encode()))
    pkg.parts["word/styles.xml"] = etree.tostring(styles, xml_declaration=True,
                                                 encoding="UTF-8", standalone=True)

    # settings: nothing needed beyond existing updateFields


if __name__ == "__main__":
    main()
