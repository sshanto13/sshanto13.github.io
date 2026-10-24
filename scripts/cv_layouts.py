"""Five single-accent-colour layouts for the user's FINAL CV content (see cv_model.py).
Content is never edited here - only presentation changes."""
import sys, copy
from pathlib import Path
import fitz
from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Pt, Emu, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
import build_cv_reviewed as B
from build_cv_reviewed import (fmt, p_border, tab_stops, cell_props, make_table, row_cant_split, fix_order,
                               mark_size, rgb, add_hyperlink)
from cv_model import parse

TMP = B.TMP
PAGE_W = 11906


class Th:
    def __init__(self, **kw):
        self.acc = "1D3B6B"
        self.ink = "111827"
        self.body = "2B2F36"
        self.mut = "5B616B"
        self.rule = "C9CDD3"
        self.tint = "EEF1F6"
        self.hf = "Arial"
        self.bf = "Calibri"
        self.bs = 9.5          # body size
        self.bl = 12.4         # body line
        self.__dict__.update(kw)


# ----------------------------------------------------------------------------- primitives
def run(p, text, size=9.5, color="2B2F36", bold=False, italic=False, font="Calibri", caps=False, sp=None, small=False):
    r = p.add_run(text)
    r.font.name = font
    rPr = r._element.get_or_add_rPr()
    rf = rPr.get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), font)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = rgb(color)
    if caps:
        r.font.all_caps = True
    if small:
        r.font.small_caps = True
    if sp:
        e = OxmlElement("w:spacing"); e.set(qn("w:val"), str(int(sp * 20))); rPr.append(e)
    return r


def P(c, before=0, after=0, line=None, align=None, keep=False, left=None, hanging=None):
    fresh = False
    if hasattr(c, "_tc"):
        ps = c.paragraphs
        fresh = len(ps) == 1 and not ps[0].runs and ps[0]._p.pPr is None
        p = ps[0] if fresh else c.add_paragraph()
    else:
        p = c.add_paragraph()
    fmt(p, before, after, line=line, align=align, keep_next=keep, left=left, hanging=hanging)
    return p


def mark_color(p, color, size=None):
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr"); pPr.append(rPr)
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    if size:
        s = OxmlElement("w:sz"); s.set(qn("w:val"), str(int(size * 2))); rPr.append(s)


def link_or_text(p, text, url, size, color, font, bold=False):
    if url:
        add_hyperlink(p, url, text, size=size, color=color, font=font)
    else:
        run(p, text, size=size, color=color, font=font, bold=bold)


def photo(ring_hex, ring_px=0, size=600):
    pdf = fitz.open(B.PHOTO_SRC_PDF)
    xref = pdf[0].get_images(full=True)[0][0]
    pix = fitz.Pixmap(pdf, xref)
    raw = TMP / "cv_photo_raw2.png"
    pix.save(str(raw))
    im = Image.open(raw).convert("RGB").resize((size, size), Image.LANCZOS)
    ring = ring_px
    out = Image.new("RGBA", (size + 2 * ring, size + 2 * ring), (0, 0, 0, 0))
    n = (size + 2 * ring) * 2
    bigr = Image.new("L", (n, n), 0)
    ImageDraw.Draw(bigr).ellipse((0, 0, n - 1, n - 1), fill=255)
    maskr = bigr.resize((size + 2 * ring, size + 2 * ring), Image.LANCZOS)
    rc = tuple(int(ring_hex[i:i + 2], 16) for i in (0, 2, 4)) + (255,)
    if ring:
        out.paste(Image.new("RGBA", out.size, rc), (0, 0), maskr)
    big = Image.new("L", (size * 2, size * 2), 0)
    ImageDraw.Draw(big).ellipse((0, 0, size * 2 - 1, size * 2 - 1), fill=255)
    mask = big.resize((size, size), Image.LANCZOS)
    out.paste(im, (ring, ring), mask)
    path = TMP / f"cv_photo_{ring_hex}_{ring_px}.png"
    out.save(path)
    return path


def new_doc(th, l=51, r=51, t=46, b=54):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = th.bf
    st.font.size = Pt(th.bs)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), th.bf)
    s = doc.sections[0]
    s.page_width, s.page_height = Emu(7560310), Emu(10692130)
    s.left_margin, s.right_margin = Pt(l), Pt(r)
    s.top_margin, s.bottom_margin = Pt(t), Pt(b)
    s.header_distance = Pt(20)
    s.footer_distance = Pt(24)
    return doc


def footer(doc, th, name, width_tw, rule=True):
    f = doc.sections[0].footer
    f.is_linked_to_previous = False
    p = f.paragraphs[0]
    fmt(p, 0, 0, line=11)
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    for pos in ("4680", "9360"):
        c = OxmlElement("w:tab"); c.set(qn("w:val"), "clear"); c.set(qn("w:pos"), pos); tabs.append(c)
    t = OxmlElement("w:tab"); t.set(qn("w:val"), "right"); t.set(qn("w:pos"), str(width_tw)); tabs.append(t)
    pPr.append(tabs)
    if rule:
        p_border(p, "top", th.rule, 4, 6)
    run(p, name, size=8, color=th.mut, font=th.bf)
    run(p, "\tPage ", size=8, color=th.mut, font=th.bf)
    B.field(p, "PAGE", 8, th.mut)
    run(p, " of ", size=8, color=th.mut, font=th.bf)
    B.field(p, "NUMPAGES", 8, th.mut)


def bullet(c, th, runs_spec, after=2.2, left=11, line=None, color=None, keep=False, size=None):
    p = c.add_paragraph(style="List Bullet") if not hasattr(c, "_tc") else c.add_paragraph(style="List Bullet")
    fmt(p, 0, after, line=line or th.bl, keep_next=keep, left=left, hanging=left)
    mark_color(p, color or th.acc, size or th.bs)
    for text, bold in runs_spec:
        run(p, text, size=size or th.bs, color=th.ink if bold else th.body, bold=bold, font=th.bf)
    return p


def role_bullets(c, th, e, after=2.2, left=11, bullet_color=None):
    out = []
    for b in e["bullets"]:
        if b[0] == "label":
            out.append(bullet(c, th, [(b[1] + " ", True), (b[2], False)], after, left, color=bullet_color))
        else:
            out.append(bullet(c, th, [(b[1], False)], after, left, color=bullet_color))
    return out


def tech_line(c, th, e, style="plain", before=2.5, left=0):
    if not e.get("tech"):
        return None
    p = P(c, before, 0, th.bl - 0.4, left=left)
    shade = th.tint if style == "tag" else None
    lab = run(p, "\u00A0Technologies:\u00A0" if shade else "Technologies: ", size=th.bs - 1, color=th.ink, bold=True)
    txt = run(p, e["tech"] + ("\u00A0" if shade else ""), size=th.bs - 1, color=th.mut)
    if shade:
        for r in (lab, txt):
            shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), shade)
            r._element.get_or_add_rPr().append(shd)
    return p


def chain(parts, head_n):
    n = len(parts)
    for i, p in enumerate(parts):
        p.paragraph_format.keep_with_next = bool((i < head_n or i == n - 2) and i < n - 1)


def set_cell_borders(c, **kw):
    cell_props(c, **kw)


def spacer(doc, pt):
    return B.spacer(doc, pt)


def card_pair_rows(doc, widths, items, build_cell, gap=None):
    """items laid out in a 2-col table: left col gets items[0::2], right col items[1::2] (reading down)."""
    t = make_table(doc, widths)
    return t


def contact_items(m):
    return m["contacts"]


# ============================================================================= A: CLASSIC
def apply_nbsp(doc, pairs):
    from docx.text.paragraph import Paragraph
    from build_cv_reviewed import join_last_two_words
    for pe in list(doc.element.body.iter(qn("w:p"))):
        para = Paragraph(pe, doc._body)
        txt = para.text.rstrip().replace("\u00a0", " ")
        for (a, b) in pairs:
            if txt.endswith(a + " " + b):
                join_last_two_words(para)


def layout_A(m, RG=1.0, HG=1.0, nbsp=()):
    th = Th(acc="1D3B6B", hf="Arial", bf="Calibri", bs=10, bl=13.2)
    doc = new_doc(th)
    W_ = 9880
    # header
    t = make_table(doc, [8250, 1630])
    L, R = t.rows[0].cells
    cell_props(L, valign="top", width=8250); cell_props(R, valign="top", width=1630)
    p = P(L, 0, 2, 31); run(p, m["name"], 27, th.ink, True, font=th.hf, sp=0.5)
    p = P(L, 0, 6, 14); run(p, m["headline"], 9.5, th.acc, True, font=th.hf, sp=0.2)
    c = m["contacts"]
    p = P(L, 0, 1.5, 13); tab_stops(p, [150, 330])
    for i, (tx, u) in enumerate(c[:3]):
        if i: run(p, "\t", 9.5)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    p = P(L, 0, 0, 13); tab_stops(p, [150, 330])
    for i, (tx, u) in enumerate(c[3:]):
        if i: run(p, "\t", 9.5)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    pr = P(R, 0, 0, None, align=WD_ALIGN_PARAGRAPH.RIGHT)
    pr.add_run().add_picture(str(photo(th.acc, 14)), width=Inches(0.92))
    p = P(doc, 4, 0, 4)
    p_border(p, "bottom", th.acc, 18, 1)
    B.mark_size(p, 1)

    def heading(text, before=12):
        p = P(doc, before * HG, 5, 14, keep=True)
        run(p, text, 10.5, th.acc, True, font=th.hf, caps=True, sp=1.3)
        p_border(p, "bottom", th.acc, 6, 3)

    H = m["headings"]
    heading(H["PROFESSIONAL SUMMARY"], 8)
    p = P(doc, 0, 5, 13.4); run(p, m["summary"], 10, th.body)
    # highlights 2x2, left accent bar
    t = make_table(doc, [4840, 200, 4840])
    hl = m["highlights"]
    for r_i in range(2):
        row = t.add_row() if r_i else t.rows[0]
        if r_i:
            for cell, w in zip(row.cells, [4840, 200, 4840]):
                cell_props(cell, width=w)
        row_cant_split(row)
        for ci, idx in ((0, r_i * 2), (2, r_i * 2 + 1)):
            cell = row.cells[ci]
            cell_props(cell, width=4840, borders={"left": (th.acc, 18)}, margins=(40, 130, 60, 40))
            p = P(cell, 0, 1, 12.5, keep=True); run(p, hl[idx][0], 9.5, th.ink, True, font=th.hf)
            p = P(cell, 0, 0, th.bl); run(p, hl[idx][1], 9.5, th.body)
        if r_i == 0:
            for cell in row.cells:
                pass
    # (row spacing via cell bottom margin)
    heading(H["TECHNICAL EXPERIENCE"], 12)
    for i, card in enumerate(m["tech_cards"]):
        parts = []
        p = P(doc, 4 if i else 0, 1.5, 13, keep=True); run(p, card["title"], 10, th.acc, True, font=th.hf); parts.append(p)
        for lab, tx in card["lines"]:
            p = P(doc, 0, 1.4, th.bl, left=0)
            if lab: run(p, lab + " ", th.bs, th.ink, True)
            run(p, tx, th.bs, th.body)
            parts.append(p)
        chain(parts, 1)

    def role(e, first=False):
        parts = []
        p = P(doc, 3 if first else 10 * RG, 0, 14, keep=True); B.tab_right(p, W_)
        long_t = len(e["title"]) + len(e["date"]) > 95
        run(p, e["title"], 10 if long_t else 10.5, th.ink, True, font=th.hf)
        run(p, "\t" + e["date"], 9, th.mut)
        parts.append(p)
        if e["company"]:
            p = P(doc, 0, 0.8 if e["intro"] else 2.5, 12.5, keep=True)
            run(p, e["company"], 9.2, th.acc, True, font=th.hf, sp=0.4); parts.append(p)
        if e["intro"]:
            p = P(doc, 0, 2.8, 12, keep=True); run(p, e["intro"], 9, th.mut, italic=True); parts.append(p)
        parts += role_bullets(doc, th, e)
        tl = tech_line(doc, th, e, "plain")
        if tl is not None: parts.append(tl)
        chain(parts, 3 + (1 if e["intro"] else 0) - (0 if e["company"] else 1))

    heading(H["PROFESSIONAL EXPERIENCE"], 12)
    for i, e in enumerate(m["experience"]): role(e, i == 0)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 14)
    for i, e in enumerate(m["consultancy"]): role(e, i == 0)
    heading(H["GOVERNMENT INITIATIVES & TRAINING"], 14)
    for e in m["gov"]: role(e, True)
    heading(H["EDUCATION"], 14)
    for e in m["education"]:
        p = P(doc, 0, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.5, th.ink, True, font=th.hf); run(p, "\t" + e["date"], 9, th.mut)
        p = P(doc, 0, 0, 12.5); run(p, e["school"], 9.2, th.acc, True, font=th.hf, sp=0.2)
    heading(H["TRAINING & PROFESSIONAL CERTIFICATIONS"], 14)
    for a, b in m["training"]:
        p = P(doc, 0, 0, 13, keep=True); run(p, a, 9.8, th.ink, True, font=th.bf)
        p = P(doc, 0, 3.5, 12.5); run(p, b, 9.3, th.mut)
    footer(doc, th, m["name"].title(), W_)
    if nbsp:
        apply_nbsp(doc, nbsp)
    return doc


# ============================================================================= B: SIDEBAR PANEL
def _layout_B_old(m):
    th = Th(acc="0F4C56", hf="Segoe UI", bf="Calibri", tint="E7F0F1", mut="5A6670")
    doc = new_doc(th, l=30, r=30, t=30, b=46)
    SW, MW = 3100, 11906 - 1200 - 3100       # sidebar, main (twips)
    t = make_table(doc, [SW, MW])
    row = t.rows[0]
    S, M = row.cells
    cell_props(S, fill=th.acc, margins=(380, 260, 300, 240), width=SW, valign="top")
    cell_props(M, margins=(300, 380, 200, 100), width=MW, valign="top")
    W_ = MW - 380 - 100
    WH = "FFFFFF"
    WS = "D5E6E8"
    # sidebar
    p = P(S, 0, 10, None, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(str(photo("FFFFFF", 12)), width=Inches(1.25))
    for tx, u in m["contacts"]:
        p = P(S, 0, 3.5, 12)
        link_or_text(p, tx, u, 9, WH, th.bf)

    def sh(text, before=14):
        p = P(S, before, 5, 12, keep=True)
        run(p, text, 8.8, WH, True, font=th.hf, caps=True, sp=1.1)
        p_border(p, "bottom", "7FA3A8", 4, 2)

    sh(m["headings"]["EDUCATION"], 16)
    for e in m["education"]:
        p = P(S, 0, 1, 12, keep=True); run(p, e["title"], 9.3, WH, True, font=th.hf)
        p = P(S, 0, 1, 11.5, keep=True); run(p, e["date"], 8.8, WS)
        p = P(S, 0, 0, 12); run(p, e["school"], 9, WS)
    sh(m["headings"]["TRAINING & PROFESSIONAL CERTIFICATIONS"])
    for a, b in m["training"]:
        p = P(S, 0, 1, 12, keep=True); run(p, a, 9.3, WH, True, font=th.hf)
        p = P(S, 0, 7, 11.8); run(p, b, 8.8, WS)
    sh(m["headings"]["GOVERNMENT INITIATIVES & TRAINING"])
    for e in m["gov"]:
        p = P(S, 0, 1, 12, keep=True); run(p, e["title"], 9.3, WH, True, font=th.hf)
        p = P(S, 0, 1.5, 11.5, keep=True); run(p, e["date"], 8.8, WS)
        p = P(S, 0, 3, 11.8, keep=True); run(p, e["intro"], 8.8, WS, italic=True)
        for b in e["bullets"]:
            q = bullet(S, th, [(b[1], False)], 0, 10, color=WH, line=11.8, size=8.8)
            for r in q.runs: r.font.color.rgb = rgb(WS)

    # main
    p = P(M, 0, 2, 33); run(p, m["name"], 27, th.acc, True, font=th.hf, sp=0.4)
    p = P(M, 0, 6, 13); run(p, m["headline"], 9, th.ink, True, font=th.hf, sp=0.3)
    p_border(p, "bottom", th.acc, 12, 5)

    def heading(text, before=12):
        p = P(M, before, 4, 14, keep=True)
        run(p, text, 10, th.acc, True, font=th.hf, caps=True, sp=1.2)

    H = m["headings"]
    heading(H["PROFESSIONAL SUMMARY"], 6)
    p = P(M, 0, 5, 13.2); run(p, m["summary"], 9.8, th.body)
    for i, (a, b) in enumerate(m["highlights"]):
        p = P(M, 0, 0, 12.6, keep=True)
        run(p, a + " ", 9.5, th.acc, True, font=th.hf); 
        p = P(M, 0, 3, th.bl); run(p, b, 9.5, th.body)
    heading(H["TECHNICAL EXPERIENCE"], 12)
    for i, card in enumerate(m["tech_cards"]):
        parts = []
        p = P(M, 5 if i else 0, 1.5, 13, keep=True); run(p, card["title"], 9.8, th.acc, True, font=th.hf); parts.append(p)
        for lab, tx in card["lines"]:
            p = P(M, 0, 1.4, th.bl)
            if lab: run(p, lab + " ", th.bs, th.ink, True)
            run(p, tx, th.bs, th.body); parts.append(p)
        chain(parts, 1)

    def role(e, first=False):
        parts = []
        p = P(M, 3 if first else 10, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.2, th.ink, True, font=th.hf)
        run(p, "\t" + e["date"], 8.6, th.mut)
        parts.append(p)
        if e["company"]:
            p = P(M, 0, 0.8 if e["intro"] else 2.5, 12.5, keep=True)
            run(p, e["company"], 9, th.acc, True, font=th.hf, sp=0.4); parts.append(p)
        if e["intro"]:
            p = P(M, 0, 2.8, 12, keep=True); run(p, e["intro"], 9, th.mut, italic=True); parts.append(p)
        parts += role_bullets(M, th, e)
        tl = tech_line(M, th, e, "tag")
        if tl is not None: parts.append(tl)
        chain(parts, 3 + (1 if e["intro"] else 0) - (0 if e["company"] else 1))

    heading(H["PROFESSIONAL EXPERIENCE"], 14)
    for i, e in enumerate(m["experience"]): role(e, i == 0)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 16)
    for i, e in enumerate(m["consultancy"]): role(e, i == 0)
    footer(doc, th, m["name"].title(), 11906 - 1200, rule=False)
    return doc


# ============================================================================= C: BANNER + CARDS
def layout_C(m):
    th = Th(acc="1F2933", ink="111111", body="2D2D2D", mut="62666B", rule="CCCCCC", tint="F1F2F3", hf="Arial", bf="Calibri")
    doc = new_doc(th, l=46, r=46, t=34, b=54)
    W_ = 11906 - 2 * 920
    t = make_table(doc, [7700, W_ - 7700])
    L, R = t.rows[0].cells
    top = {"top": (th.acc, 48)}
    cell_props(L, fill=th.tint, borders=top, margins=(260, 380, 240, 100), width=7700, valign="center")
    cell_props(R, fill=th.tint, borders=top, margins=(200, 100, 200, 300), width=W_ - 7700, valign="center")
    p = P(L, 0, 3, 32); run(p, m["name"], 28, th.ink, True, font=th.hf, sp=0.6)
    p = P(L, 0, 8, 13); run(p, m["headline"], 8.8, th.acc, True, font=th.hf, sp=0.3)
    c = m["contacts"]
    p = P(L, 0, 1.5, 12.6); tab_stops(p, [130, 255])
    for i, (tx, u) in enumerate(c[:3]):
        if i: run(p, "\t", 9)
        link_or_text(p, tx, u, 9.2, th.ink, th.bf)
    p = P(L, 0, 0, 12.6); tab_stops(p, [130, 255])
    for i, (tx, u) in enumerate(c[3:]):
        if i: run(p, "\t", 9)
        link_or_text(p, tx, u, 9.2, th.ink, th.bf)
    p = P(R, 0, 0, None, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p.add_run().add_picture(str(photo(th.acc, 14)), width=Inches(0.95))
    row_cant_split(t.rows[0])

    def heading(text, before=13):
        p = P(doc, before, 5, 14, keep=True)
        p_border(p, "left", th.acc, 36, 6)
        run(p, text, 10.5, th.ink, True, font=th.hf, caps=True, sp=1.2)

    H = m["headings"]
    P(doc, 0, 0, 5)
    heading(H["PROFESSIONAL SUMMARY"], 4)
    p = P(doc, 0, 6, 13.4); run(p, m["summary"], 10, th.body)
    # 4 highlight cells with accent top rule
    cw = [2410, 2410, 2410, 2410]
    gaps = (W_ - sum(cw)) // 3
    t = make_table(doc, [cw[0], gaps, cw[1], gaps, cw[2], gaps, cw[3]])
    row = t.rows[0]; row_cant_split(row)
    for i, cell in enumerate(row.cells):
        if i % 2:
            cell_props(cell, width=gaps); continue
        a, b = m["highlights"][i // 2]
        cell_props(cell, width=cw[0], borders={"top": (th.acc, 18)}, fill=th.tint, margins=(90, 110, 90, 100))
        p = P(cell, 0, 2, 12.4); run(p, a, 9.3, th.ink, True, font=th.hf)
        p = P(cell, 0, 0, 12); run(p, b, 8.8, th.body)
    heading(H["TECHNICAL EXPERIENCE"], 14)
    # 2-col cards (reading order: left column 1,2,3 then right column 4,5 is hard to scan, so use row pairs)
    cards = m["tech_cards"]
    t = make_table(doc, [4860, 160, 4860])
    first = True
    for i in range(0, len(cards), 2):
        row = t.rows[0] if first else t.add_row()
        if not first:
            for cell, w in zip(row.cells, [4860, 160, 4860]): cell_props(cell, width=w)
        first = False
        row_cant_split(row)
        full = i + 1 >= len(cards)
        if full:
            merged = row.cells[0].merge(row.cells[2])
        for ci, k in ((0, i), (2, i + 1)):
            if k >= len(cards): continue
            card = cards[k]
            cell = merged if full else row.cells[ci]
            cell_props(cell, width=(W_ if full else 4860), borders={"left": (th.acc, 18), "bottom": (th.rule, 4)}, margins=(40, 130, 90, 60))
            p = P(cell, 0, 2, 12.8); run(p, card["title"], 9.5, th.ink, True, font=th.hf)
            for lab, tx in card["lines"]:
                p = P(cell, 0, 1.6, 11.8)
                if lab: run(p, lab + " ", 9, th.ink, True)
                run(p, tx, 9, th.body)
    def role(e, first=False):
        parts = []
        p = P(doc, 4 if first else 10, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.5, th.ink, True, font=th.hf)
        run(p, "\t" + e["date"], 9, th.mut)
        parts.append(p)
        if e["company"]:
            p = P(doc, 0, 0.8 if e["intro"] else 2.5, 12.5, keep=True)
            run(p, e["company"], 9.2, th.ink, True, font=th.hf, sp=0.6); parts.append(p)
            p_border(p, "left", th.acc, 18, 5)
        if e["intro"]:
            p = P(doc, 0, 2.8, 12, keep=True); run(p, e["intro"], 9, th.mut, italic=True); parts.append(p)
        parts += role_bullets(doc, th, e, bullet_color=th.ink)
        tl = tech_line(doc, th, e, "tag")
        if tl is not None: parts.append(tl)
        chain(parts, 3 + (1 if e["intro"] else 0) - (0 if e["company"] else 1))

    heading(H["PROFESSIONAL EXPERIENCE"], 14)
    for i, e in enumerate(m["experience"]): role(e, i == 0)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 14)
    for i, e in enumerate(m["consultancy"]): role(e, i == 0)
    heading(H["GOVERNMENT INITIATIVES & TRAINING"], 14)
    for e in m["gov"]: role(e, True)
    # education + training side by side
    heading("", 0) if False else None
    t = make_table(doc, [4860, 160, 4860])
    row = t.rows[0]; row_cant_split(row)
    A, _, Cc = row.cells
    p = P(doc, 0, 0, 1)  # placeholder keeps structure simple
    doc.element.body.remove(p._p)
    cell_props(A, width=4860); cell_props(Cc, width=4860)
    p = P(A, 10, 4, 14, keep=True); p_border(p, "left", th.acc, 36, 6)
    run(p, H["EDUCATION"], 10.5, th.ink, True, font=th.hf, caps=True, sp=1.2)
    for e in m["education"]:
        p = P(A, 0, 1, 13.5, keep=True); run(p, e["title"], 10, th.ink, True, font=th.hf)
        p = P(A, 0, 1, 12, keep=True); run(p, e["date"], 9, th.mut)
        p = P(A, 0, 0, 12.5); run(p, e["school"], 9.2, th.ink, True, font=th.hf, sp=0.2)
    p = P(Cc, 10, 4, 14, keep=True); p_border(p, "left", th.acc, 36, 6)
    run(p, H["TRAINING & PROFESSIONAL CERTIFICATIONS"], 10.5, th.ink, True, font=th.hf, caps=True, sp=0.6)
    for a, b in m["training"]:
        p = P(Cc, 0, 0, 13, keep=True); run(p, a, 9.8, th.ink, True)
        p = P(Cc, 0, 3.5, 12); run(p, b, 9.2, th.mut)
    footer(doc, th, m["name"].title(), W_)
    return doc


# ============================================================================= D: TIMELINE RAIL
def layout_D(m):
    th = Th(acc="7A1F35", ink="1B1B1F", body="2E2E33", mut="63636B", rule="D6D0D2", tint="F6EFF1", hf="Georgia", bf="Calibri")
    doc = new_doc(th, l=50, r=50, t=44, b=54)
    W_ = 9906 - 26
    RAIL, GAP = 2050, 0
    RW = W_ - RAIL
    # header
    t = make_table(doc, [8350, W_ - 8350])
    L, R = t.rows[0].cells
    cell_props(L, valign="top", width=8350); cell_props(R, valign="top", width=W_ - 8350)
    p = P(L, 0, 2, 34); run(p, m["name"], 30, th.acc, True, font=th.hf, sp=0.2)
    p = P(L, 0, 7, 13); run(p, m["headline"], 8.8, th.ink, True, font="Arial", sp=0.9)
    c = m["contacts"]
    p = P(L, 0, 1.5, 12.8); tab_stops(p, [150, 330])
    for i, (tx, u) in enumerate(c[:3]):
        if i: run(p, "\t", 9.5)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    p = P(L, 0, 0, 12.8); tab_stops(p, [150, 330])
    for i, (tx, u) in enumerate(c[3:]):
        if i: run(p, "\t", 9.5)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    p = P(R, 0, 0, None, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p.add_run().add_picture(str(photo(th.acc, 0)), width=Inches(0.95))

    H = m["headings"]

    def heading(text, before=14):
        p = P(doc, before, 6, 16, keep=True)
        p_border(p, "top", th.acc, 12, 5)
        run(p, text, 11.5, th.acc, True, font=th.hf, caps=True, sp=1.0)

    def rail_table(rows, keep_rows=True):
        """rows: list of (rail_fn, body_fn). returns table"""
        t = make_table(doc, [RAIL, RW])
        for i, (rf, bf) in enumerate(rows):
            row = t.rows[0] if i == 0 else t.add_row()
            a, b = row.cells
            cell_props(a, width=RAIL, margins=(0, 0, 0, 200))
            cell_props(b, width=RW, borders={"left": (th.acc, 8)}, margins=(0, 220, 150 if i < len(rows) - 1 else 0, 0))
            if keep_rows: row_cant_split(row)
            rf(a); bf(b)
        return t

    def rail_text(txt, bold=False, color=None, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT, font=None, before=0):
        def f(cell):
            p = P(cell, before, 0, 11.6, align=align)
            run(p, txt, size, color or th.mut, bold, font=font or th.bf)
        return f

    heading(H["PROFESSIONAL SUMMARY"], 10)
    p = P(doc, 0, 6, 13.6); run(p, m["summary"], 10, th.body)
    rows = []
    for a, b in m["highlights"]:
        def rf(cell, a=a):
            p = P(cell, 0, 0, 12, align=WD_ALIGN_PARAGRAPH.RIGHT); run(p, a, 9.2, th.acc, True, font="Arial")
        def bf(cell, b=b):
            p = P(cell, 0, 0, th.bl); run(p, b, 9.5, th.body)
        rows.append((rf, bf))
    rail_table(rows)

    heading(H["TECHNICAL EXPERIENCE"], 14)
    rows = []
    for card in m["tech_cards"]:
        def rf(cell, card=card):
            p = P(cell, 0, 0, 12, align=WD_ALIGN_PARAGRAPH.RIGHT); run(p, card["title"], 9.2, th.acc, True, font="Arial")
        def bf(cell, card=card):
            for lab, tx in card["lines"]:
                p = P(cell, 0, 1.6, th.bl)
                if lab: run(p, lab + " ", th.bs, th.ink, True)
                run(p, tx, th.bs, th.body)
        rows.append((rf, bf))
    rail_table(rows)

    def role_rows(items):
        rows = []
        for e in items:
            def rf(cell, e=e):
                p = P(cell, 0, 0, 12, align=WD_ALIGN_PARAGRAPH.RIGHT); run(p, e["date"], 8.8, th.acc, True, font=th.bf)
            def bf(cell, e=e):
                p = P(cell, 0, 0, 14); run(p, e["title"], 10.5, th.ink, True, font=th.hf)
                if e["company"]:
                    p = P(cell, 0, 0.8 if e["intro"] else 2.5, 12.4); run(p, e["company"], 9, th.acc, True, font="Arial", sp=0.4)
                if e["intro"]:
                    p = P(cell, 0, 2.8, 12); run(p, e["intro"], 9, th.mut, italic=True)
                role_bullets(cell, th, e, after=2.0)
                tech_line(cell, th, e, "tag")
            rows.append((rf, bf))
        return rows

    heading(H["PROFESSIONAL EXPERIENCE"], 14)
    rail_table(role_rows(m["experience"]), keep_rows=False)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 14)
    rail_table(role_rows(m["consultancy"]), keep_rows=True)
    heading(H["GOVERNMENT INITIATIVES & TRAINING"], 14)
    rail_table(role_rows(m["gov"]))
    heading(H["EDUCATION"], 14)
    rows = []
    for e in m["education"]:
        rows.append((rail_text(e["date"], True, th.acc, 8.8),
                     lambda cell, e=e: (run(P(cell, 0, 0, 14), e["title"], 10.5, th.ink, True, font=th.hf),
                                        run(P(cell, 0, 0, 12.4), e["school"], 9.2, th.acc, True, font="Arial", sp=0.2))))
    rail_table(rows)
    heading(H["TRAINING & PROFESSIONAL CERTIFICATIONS"], 14)
    rows = []
    for a, b in m["training"]:
        rows.append((rail_text(a, True, th.acc, 9.0, font="Arial"),
                     lambda cell, b=b: run(P(cell, 0, 0, 12.4), b, 9.5, th.body)))
    rail_table(rows)
    footer(doc, th, m["name"].title(), W_)
    return doc


# ============================================================================= E: EDITORIAL (centred, serif)
def layout_E(m):
    th = Th(acc="1F5B3D", ink="14181A", body="2A2F30", mut="5E6566", rule="C8CFCB", tint="EFF4F1", hf="Cambria", bf="Cambria", bs=9.4, bl=12.4)
    doc = new_doc(th, l=54, r=54, t=44, b=54)
    W_ = 11906 - 2 * 1080
    t = make_table(doc, [1500, W_ - 3000, 1500])
    A, Bc, C_ = t.rows[0].cells
    cell_props(A, width=1500, valign="center"); cell_props(Bc, width=W_ - 3000, valign="center"); cell_props(C_, width=1500)
    p = P(A, 0, 0, None); p.add_run().add_picture(str(photo(th.acc, 12)), width=Inches(0.92))
    p = P(Bc, 0, 2, 34, align=WD_ALIGN_PARAGRAPH.CENTER); run(p, m["name"], 28, th.ink, True, font=th.hf, sp=3)
    p = P(Bc, 0, 6, 13, align=WD_ALIGN_PARAGRAPH.CENTER); run(p, m["headline"], 8.6, th.acc, True, font="Arial", sp=0.8)
    c = m["contacts"]
    p = P(Bc, 0, 1.5, 12.6, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (tx, u) in enumerate(c[:3]):
        if i: run(p, "     ", 9.5, font=th.bf)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    p = P(Bc, 0, 0, 12.6, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (tx, u) in enumerate(c[3:]):
        if i: run(p, "     ", 9.5, font=th.bf)
        link_or_text(p, tx, u, 9.5, th.ink, th.bf)
    p = P(doc, 5, 0, 4); p_border(p, "bottom", th.acc, 12, 1); B.mark_size(p, 1)

    H = m["headings"]

    def heading(text, before=12):
        p = P(doc, before, 5, 15, align=WD_ALIGN_PARAGRAPH.CENTER, keep=True)
        run(p, text, 10.5, th.acc, True, font=th.hf, caps=True, sp=2.0)
        p_border(p, "bottom", th.rule, 4, 3)

    heading(H["PROFESSIONAL SUMMARY"], 8)
    p = P(doc, 0, 6, 13.2, align=WD_ALIGN_PARAGRAPH.JUSTIFY); run(p, m["summary"], 9.8, th.body, font=th.bf)
    cw = (W_ - 3 * 100) // 4
    t = make_table(doc, [cw, 100, cw, 100, cw, 100, W_ - 3 * cw - 3 * 100])
    row = t.rows[0]; row_cant_split(row)
    for i, cell in enumerate(row.cells):
        if i % 2:
            cell_props(cell, width=100, borders={"left": (th.rule, 4)}) if False else cell_props(cell, width=100)
            continue
        a, b = m["highlights"][i // 2]
        bd = {"left": (th.rule, 4)} if i else None
        cell_props(cell, width=cw, borders=bd, margins=(0, 100, 0, 40))
        p = P(cell, 0, 2, 12.4); run(p, a, 9.3, th.acc, True, font=th.hf)
        p = P(cell, 0, 0, 11.8); run(p, b, 8.7, th.body, font=th.bf)

    heading(H["TECHNICAL EXPERIENCE"], 14)
    cards = m["tech_cards"]
    half = (W_ - 240) // 2
    t = make_table(doc, [half, 240, half])
    first = True
    for i in range(0, len(cards), 2):
        row = t.rows[0] if first else t.add_row()
        if not first:
            for cell, w in zip(row.cells, [half, 240, half]): cell_props(cell, width=w)
        first = False
        row_cant_split(row)
        full = i + 1 >= len(cards)
        if full:
            merged = row.cells[0].merge(row.cells[2])
        for ci, k in ((0, i), (2, i + 1)):
            if k >= len(cards): continue
            card = cards[k]
            cell = merged if full else row.cells[ci]
            cell_props(cell, width=(W_ if full else half), margins=(0, 0, 110, 0))
            p = P(cell, 0, 1.5, 12.8); run(p, card["title"], 9.6, th.acc, True, font=th.hf)
            p_border(p, "bottom", th.rule, 4, 2)
            for lab, tx in card["lines"]:
                p = P(cell, 0, 1.6, 11.8)
                if lab: run(p, lab + " ", 8.9, th.ink, True, font=th.bf)
                run(p, tx, 8.9, th.body, font=th.bf)

    def role(e, first=False):
        parts = []
        p = P(doc, 3 if first else 10, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.6, th.ink, True, font=th.hf)
        run(p, "\t" + e["date"], 9, th.mut, italic=True, font=th.bf)
        parts.append(p)
        if e["company"]:
            p = P(doc, 0, 0.8 if e["intro"] else 2.5, 12.4, keep=True)
            run(p, e["company"], 9.4, th.acc, True, font=th.hf, sp=0.5); parts.append(p)
        if e["intro"]:
            p = P(doc, 0, 2.8, 12, keep=True); run(p, e["intro"], 9, th.mut, italic=True, font=th.bf); parts.append(p)
        parts += role_bullets(doc, th, e)
        tl = tech_line(doc, th, e, "plain")
        if tl is not None: parts.append(tl)
        chain(parts, 3 + (1 if e["intro"] else 0) - (0 if e["company"] else 1))

    heading(H["PROFESSIONAL EXPERIENCE"], 14)
    for i, e in enumerate(m["experience"]): role(e, i == 0)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 14)
    for i, e in enumerate(m["consultancy"]): role(e, i == 0)
    heading(H["GOVERNMENT INITIATIVES & TRAINING"], 14)
    for e in m["gov"]: role(e, True)
    heading(H["EDUCATION"], 14)
    for e in m["education"]:
        p = P(doc, 0, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.6, th.ink, True, font=th.hf); run(p, "\t" + e["date"], 9, th.mut, italic=True)
        p = P(doc, 0, 0, 12.4); run(p, e["school"], 9.4, th.acc, True, font=th.hf, sp=0.4)
    heading(H["TRAINING & PROFESSIONAL CERTIFICATIONS"], 14)
    for a, b in m["training"]:
        p = P(doc, 0, 0, 13, keep=True); run(p, a, 9.8, th.ink, True, font=th.hf)
        p = P(doc, 0, 3.5, 12.4); run(p, b, 9.3, th.mut, italic=True)
    footer(doc, th, m["name"].title(), W_)
    return doc


def layout_B(m):
    """Page-1 panel (light teal sidebar + main) then full-width flow."""
    th = Th(acc="0F5560", hf="Segoe UI", bf="Calibri", tint="E2EEEF", mut="55646A", rule="BCD0D3")
    doc = new_doc(th, l=51, r=51, t=40, b=54)
    W_ = 9866
    SW, GAP = 2950, 0
    MW = W_ - SW
    t = make_table(doc, [SW, MW])
    S, M = t.rows[0].cells
    cell_props(S, fill=th.tint, margins=(300, 230, 260, 200), width=SW, valign="top",
               borders={"top": (th.acc, 36)})
    cell_props(M, margins=(60, 360, 100, 0), width=MW, valign="top")
    H = m["headings"]
    p = P(S, 0, 10, None, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(str(photo(th.acc, 14)), width=Inches(1.15))
    for tx, u in m["contacts"]:
        p = P(S, 0, 3.4, 12.2)
        link_or_text(p, tx, u, 9.3, th.ink, th.bf)

    def sh(text, before=14):
        p = P(S, before, 4, 12, keep=False)
        run(p, text, 8.6, th.acc, True, font=th.hf, caps=True, sp=1.0)
        p_border(p, "bottom", th.acc, 6, 2)

    sh(H["EDUCATION"], 12)
    for e in m["education"]:
        p = P(S, 0, 1, 12.2); run(p, e["title"], 9.2, th.ink, True, font=th.hf)
        p = P(S, 0, 1, 11.6); run(p, e["date"], 8.8, th.mut)
        p = P(S, 0, 0, 12); run(p, e["school"], 9, th.acc, True, font=th.hf)
    sh(H["TRAINING & PROFESSIONAL CERTIFICATIONS"], 12)
    for a, b in m["training"]:
        p = P(S, 0, 1, 12.2); run(p, a, 9.2, th.ink, True, font=th.hf)
        p = P(S, 0, 6, 11.8); run(p, b, 8.8, th.mut)

    p = P(M, 0, 2, 33); run(p, m["name"], 27, th.acc, True, font=th.hf, sp=0.4)
    p = P(M, 0, 6, 13); run(p, m["headline"], 8.8, th.ink, True, font=th.hf, sp=0.3)
    p = P(M, 0, 5, 13); run(p, H["PROFESSIONAL SUMMARY"], 9.6, th.acc, True, font=th.hf, caps=True, sp=1.2)
    p_border(p, "top", th.rule, 4, 5)
    p = P(M, 0, 6, 13.2); run(p, m["summary"], 9.8, th.body)
    for a, b in m["highlights"]:
        p = P(M, 0, 0.5, 12.4); run(p, a, 9.5, th.acc, True, font=th.hf)
        p = P(M, 0, 4, th.bl); run(p, b, 9.5, th.body)
    row_cant_split(t.rows[0])

    def heading(text, before=14):
        p = P(doc, before, 5, 15, keep=True)
        p.paragraph_format.keep_together = True
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), th.tint)
        pPr.append(shd)
        p_border(p, "left", th.acc, 36, 5)
        run(p, text, 10, th.acc, True, font=th.hf, caps=True, sp=1.2)

    heading(H["TECHNICAL EXPERIENCE"], 14)
    for i, card in enumerate(m["tech_cards"]):
        parts = []
        p = P(doc, 4 if i else 0, 1.5, 13, keep=True); run(p, card["title"], 9.8, th.acc, True, font=th.hf); parts.append(p)
        for lab, tx in card["lines"]:
            p = P(doc, 0, 1.4, th.bl)
            if lab: run(p, lab + " ", th.bs, th.ink, True)
            run(p, tx, th.bs, th.body); parts.append(p)
        chain(parts, 1)

    def role(e, first=False):
        parts = []
        p = P(doc, 3 if first else 10, 0, 14, keep=True); B.tab_right(p, W_)
        run(p, e["title"], 10.3, th.ink, True, font=th.hf)
        run(p, "\t" + e["date"], 8.8, th.mut)
        parts.append(p)
        if e["company"]:
            p = P(doc, 0, 0.8 if e["intro"] else 2.5, 12.5, keep=True)
            run(p, e["company"], 9, th.acc, True, font=th.hf, sp=0.4); parts.append(p)
        if e["intro"]:
            p = P(doc, 0, 2.8, 12, keep=True); run(p, e["intro"], 9, th.mut, italic=True); parts.append(p)
        parts += role_bullets(doc, th, e)
        tl = tech_line(doc, th, e, "tag")
        if tl is not None: parts.append(tl)
        chain(parts, 3 + (1 if e["intro"] else 0) - (0 if e["company"] else 1))

    heading(H["PROFESSIONAL EXPERIENCE"], 14)
    for i, e in enumerate(m["experience"]): role(e, i == 0)
    heading(H["GLOBAL PROJECT ENGAGEMENTS"], 14)
    for i, e in enumerate(m["consultancy"]): role(e, i == 0)
    heading(H["GOVERNMENT INITIATIVES & TRAINING"], 14)
    for e in m["gov"]: role(e, True)
    footer(doc, th, m["name"].title(), W_)
    return doc


LAYOUTS = {
    "A_Classic_Navy": layout_A,
    "B_Sidebar_Teal": layout_B,
    "C_Banner_Mono": layout_C,
    "D_Timeline_Burgundy": layout_D,
    "E_Editorial_Green": layout_E,
}
