"""Typography-only fix for Sumon_Ray_ACI_Gemini_CV.docx (no wording changes).

* gives un-coloured runs the same slate colour as the surrounding text
* removes stray empty paragraphs so gaps are equal everywhere
* formats the 'TRAINING & PROFESSIONAL CERTIFICATIONS' heading like the others
* aligns the 2nd contact row with a tab stop instead of literal spaces
* trims trailing spaces in headings
* removes single-word last lines (widows) with a tiny negative tracking
* spreads spacing so all three pages fill down to the footer
"""
import copy
import os
import re
import sys
import tempfile
from pathlib import Path

import fitz
import win32com.client as w
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
SRC = BASE / "Sumon_Ray_ACI_Gemini_CV.docx"
OUT = BASE / "Sumon_Ray_ACI_Gemini_CV_typography.docx"
TMP = Path(tempfile.gettempdir())

BOTTOM_LIMIT = 774
SLATE_BODY = "334155"
MAX_TRACK = -0.45

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]


def reorder(el, order):
    def key(c):
        n = c.tag.split("}")[1]
        return order.index(n) if n in order else len(order)
    kids = list(el)
    for k in kids:
        el.remove(k)
    for k in sorted(kids, key=key):
        el.append(k)


def fix_order(root):
    for tag, order in (("w:pPr", PPR_ORDER), ("w:rPr", RPR_ORDER), ("w:tcPr", TCPR_ORDER)):
        for el in root.iter(qn(tag)):
            reorder(el, order)


def para_text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def line_pt(para):
    pf = para.paragraph_format
    ln = pf.line_spacing
    return ln.pt if hasattr(ln, "pt") else None


# ---------------------------------------------------------------- base fixes
def base_fixes(d):
    P = lambda el: Paragraph(el, d._body)
    body = list(d.element.body)

    # 1. colour for runs that have none
    for p in d.element.body.iter(qn("w:p")):
        for r in P(p).runs:
            if r.text.strip() and (r.font.color is None or r.font.color.type is None):
                r.font.color.rgb = RGBColor.from_string(SLATE_BODY)

    # 2a. stray empty paragraphs between blocks (keep thin designed spacers + sectPr carrier)
    last_p = [e for e in body if e.tag == qn("w:p")][-1]
    for el in body[:-1]:
        if el.tag != qn("w:p") or el is last_p or para_text(el).strip():
            continue
        if el.pPr is not None and el.pPr.find(qn("w:sectPr")) is not None:
            continue
        ln = line_pt(P(el))
        if ln is not None and ln <= 8:
            continue
        el.getparent().remove(el)

    # 2b. empty paragraphs inside table cells that also hold text
    for tc in d.element.body.iter(qn("w:tc")):
        ps = tc.findall(qn("w:p"))
        if any(para_text(p).strip() for p in ps):
            for p in ps:
                if not para_text(p).strip():
                    tc.remove(p)

    # the compulsory last paragraph must not push a 4th page
    pf = P(last_p).paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(1)
    pPr = last_p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    for old in rPr.findall(qn("w:sz")):
        rPr.remove(old)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "2")
    rPr.append(sz)

    # 3. TRAINING heading formatted like the EDUCATION heading
    paras = [P(e) for e in d.element.body if e.tag == qn("w:p")]
    edu = next(p for p in paras if p.text.strip() == "EDUCATION")
    train = next(p for p in paras if p.text.strip().startswith("TRAINING & PROFESSIONAL"))
    new_pPr = copy.deepcopy(edu._p.pPr)
    if train._p.pPr is not None:
        train._p.remove(train._p.pPr)
    train._p.insert(0, new_pPr)
    ref_rPr = copy.deepcopy(edu.runs[0]._r.rPr)
    text = train.text.strip()
    for r in list(train.runs):
        r._r.getparent().remove(r._r)
    run = train.add_run(text)
    if run._r.rPr is not None:
        run._r.remove(run._r.rPr)
    run._r.insert(0, ref_rPr)

    # spacers around that heading: none above (heading has its own space), one (the 4pt one) below
    prev = train._p.getprevious()
    if prev is not None and prev.tag == qn("w:p") and not para_text(prev).strip():
        prev.getparent().remove(prev)
    n1 = train._p.getnext()
    n2 = n1.getnext() if n1 is not None else None
    if (n1 is not None and n2 is not None and n1.tag == qn("w:p") and n2.tag == qn("w:p")
            and not para_text(n1).strip() and not para_text(n2).strip()):
        n2.getparent().remove(n2)

    # 4. trailing spaces in headings
    for p in paras:
        if p.runs and p.text.isupper() and p.runs[-1].text.endswith(" "):
            p.runs[-1].text = p.runs[-1].text.rstrip(" ")

    # 5. contact row 2: literal spaces -> tab stop
    hdr = d.tables[0].rows[0].cells[0]
    for p in hdr.paragraphs:
        joined = "".join(r.text for r in p.runs)
        if "github.com/sshanto13" in joined and "    " in joined:
            for r in p.runs:
                if "    " in r.text:
                    r.text = r.text.replace("    ", "\t")
            pPr = p._p.get_or_add_pPr()
            tabs = pPr.find(qn("w:tabs"))
            if tabs is None:
                tabs = OxmlElement("w:tabs")
                pPr.append(tabs)
            t = OxmlElement("w:tab")
            t.set(qn("w:val"), "left")
            t.set(qn("w:pos"), str(126 * 20))
            tabs.append(t)


# ---------------------------------------------------------------- scaling
def ranges(d):
    body = list(d.element.body)
    idx = lambda txt: next(i for i, e in enumerate(body) if para_text(e).strip() == txt)
    sect = next(i for i, e in enumerate(body) if e.tag == qn("w:p") and e.pPr is not None
                and e.pPr.find(qn("w:sectPr")) is not None)
    s_sum = idx("PROFESSIONAL SUMMARY")
    s_exp = idx("PROFESSIONAL EXPERIENCE")
    g_idx = idx("GLOBAL & LOCAL CONSULTANCY ENGAGEMENTS")
    return {
        "p1": (s_sum, s_exp - 1),
        "p2": (s_exp, sect),
        "p3a": (sect + 1, g_idx - 1),               # Dhaka Web Host block: same rhythm as page 2
        "p3b": (g_idx, len(body) - 2),
    }


def scale_para(p, k, g, orig):
    pf = p.paragraph_format
    key = p._p                      # the element itself (kept alive by the dict) - never id()
    if key not in orig:
        ln = pf.line_spacing
        orig[key] = (
            ln.pt if (pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY and hasattr(ln, "pt")) else None,
            pf.space_before.pt if pf.space_before is not None else 0.0,
            pf.space_after.pt if pf.space_after is not None else 0.0,
        )
    line0, b0, a0 = orig[key]
    if line0 is not None:
        thin = line0 <= 8 and not p.text.strip()
        pf.line_spacing = Pt(line0 * (g if thin else k))
    pf.space_before = Pt(b0 * g)
    pf.space_after = Pt(a0 * g)


def scale_range(d, lo, hi, t):
    k = 1 + 0.12 * t
    g = 1 + 1.6 * t
    body = list(d.element.body)
    orig = {}
    for el in body[lo:hi + 1]:
        if el.tag == qn("w:p"):
            scale_para(Paragraph(el, d._body), k, g, orig)
        elif el.tag == qn("w:tbl"):
            for tc in el.iter(qn("w:tc")):
                tcPr = tc.find(qn("w:tcPr"))
                mar = tcPr.find(qn("w:tcMar")) if tcPr is not None else None
                if mar is not None:
                    for edge in ("top", "bottom"):
                        e = mar.find(qn(f"w:{edge}"))
                        if e is not None:
                            e.set(qn("w:w"), str(int(float(e.get(qn("w:w"))) * (1 + 0.9 * t))))
                for pe in tc.findall(qn("w:p")):
                    scale_para(Paragraph(pe, d._body), k, 1 + 1.0 * t, orig)


def join_last_two_words(para):
    """Non-breaking space between the last two words so the last line never holds a single word."""
    text = para.text
    pos = len(text.rstrip()) - 1
    while pos >= 0 and text[pos] != " ":
        pos -= 1
    if pos < 0:
        return
    off = 0
    for r in para.runs:
        n = len(r.text)
        if off <= pos < off + n:
            s = r.text
            i = pos - off
            r.text = s[:i] + "\u00a0" + s[i + 1:]
            return
        off += n


# ---------------------------------------------------------------- measuring
class Measurer:
    def __init__(self):
        self.app = w.Dispatch("Word.Application")
        self.app.Visible = False
        self.n = 0

    def measure(self, d):
        fix_order(d.element)
        for s in d.sections:
            fix_order(s.footer._element)
        self.n += 1
        docx = TMP / f"_fx{self.n}.docx"
        pdfp = TMP / f"_fx{self.n}.pdf"
        d.save(docx)
        doc = self.app.Documents.Open(str(docx), ReadOnly=True)
        pages = doc.ComputeStatistics(2)
        doc.SaveAs2(str(pdfp), 17)
        doc.Close(False)
        pdf = fitz.open(pdfp)
        info = []
        for pg in pdf:
            lines = []
            for b in pg.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    txt = "".join(s["text"] for s in l["spans"]).strip()
                    if txt:
                        lines.append((l["bbox"][1], l["bbox"][3], l["bbox"][0], l["bbox"][2], txt))
            lines.sort()
            bottom = max((ln[1] for ln in lines if ln[0] < 792), default=0)
            widows = []
            for i in range(1, len(lines)):
                y0, y1, x0, x1, txt = lines[i]
                py0, py1, px0, px1, ptxt = lines[i - 1]
                if (y0 < 792 and len(txt.split()) == 1 and len(txt) > 3 and 56 < x0 < 80
                        and 0 < y0 - py0 < 16 and px1 > 460 and abs(px0 - x0) < 14):
                    widows.append((ptxt.split()[-1], txt))
            info.append((bottom, widows))
        pdf.close()
        for f in (docx, pdfp):
            try:
                os.remove(f)
            except OSError:
                pass
        return pages, info

    def close(self):
        self.app.Quit()


def build(t, tracks):
    d = Document(SRC)
    base_fixes(d)
    for p in [Paragraph(e, d._body) for e in d.element.body.iter(qn("w:p"))]:
        txt = p.text.strip()
        for (prevw, wd) in tracks:
            if txt.endswith(prevw + " " + wd):
                join_last_two_words(p)
    r = ranges(d)
    scale_range(d, *r["p1"], t["p1"])
    scale_range(d, *r["p2"], t["p2"])
    scale_range(d, *r["p3a"], t["p2"])
    scale_range(d, *r["p3b"], t["p3"])
    return d


def solve(m, key, page_idx, t, tracks):
    lo, hi = 0.0, 1.6
    best = 0.0
    for _ in range(7):
        mid = (lo + hi) / 2
        tt = dict(t)
        tt[key] = mid
        pages, info = m.measure(build(tt, tracks))
        if pages == 3 and info[page_idx][0] <= BOTTOM_LIMIT:
            best, lo = mid, mid
        else:
            hi = mid
    return best


def main():
    m = Measurer()
    tracks = {}
    t = {"p1": 0.0, "p2": 0.0, "p3": 0.0}
    try:
        for rnd in range(4):
            pages, info = m.measure(build(t, tracks))
            widows = [x for pg in info for x in pg[1]]
            print("round", rnd, "pages", pages, "bottoms", [round(i[0]) for i in info], "widows", widows)
            changed = False
            for wd in widows:
                if wd not in tracks:
                    tracks[wd] = True
                    changed = True
            if not widows or not changed:
                break
        t["p2"] = solve(m, "p2", 1, t, tracks)
        t["p1"] = solve(m, "p1", 0, t, tracks)
        t["p3"] = solve(m, "p3", 2, t, tracks)
        d = build(t, tracks)
        pages, info = m.measure(d)
        print("final t", {k: round(v, 3) for k, v in t.items()}, "pages", pages,
              "bottoms", [round(i[0]) for i in info], "widows left", [x for pg in info for x in pg[1]],
              "tracking", tracks)
        fix_order(d.element)
        for s in d.sections:
            fix_order(s.footer._element)
        d.save(OUT)
        print("saved", OUT)
    finally:
        m.close()


if __name__ == "__main__":
    main()
