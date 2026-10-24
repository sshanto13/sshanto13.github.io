"""Convert the Gemini-made 'Sumon Ray- ACI.pdf' CV into an editable Word file,
reproducing its layout (header + round photo, summary box, cards, tags, footer)."""
import os
import tempfile
from pathlib import Path

import fitz
from PIL import Image, ImageDraw
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, RGBColor, Emu

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
PDF = ROOT / "Sumon Ray- ACI.pdf"
OUT = ROOT / "Sumon_Ray_ACI_Final_Clean.docx"

FONT = "Segoe UI"
EMOJI = "Segoe UI Emoji"

DARK = "0F172A"
INK = "1E293B"
BODY = "334155"
SLATE = "475569"
MUTED = "64748B"
BLUE = "2563EB"
NAVY = "1E40AF"
RULE = "CBD5E1"
CARD = "E2E8F0"
BG = "F8FAFC"
TAG = "F1F5F9"

CONTENT_TW = 9880          # 494pt in twips
TOP = 52


# ----------------------------------------------------------------- low-level
def rgb(h):
    return RGBColor.from_string(h)


def run(p, text, size=8.5, color=BODY, bold=False, italic=False, font=FONT, shade=None, spacing=None):
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
    if shade:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), shade)
        rPr.append(shd)
    if spacing:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        rPr.append(sp)
    return r


def fmt(p, before=0, after=0, line=None, align=None, keep_next=False, left=None, hanging=None, keep_lines=True):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line)
    if align is not None:
        p.alignment = align
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    if left is not None:
        pf.left_indent = Pt(left)
    if hanging is not None:
        pf.first_line_indent = Pt(-hanging)
    return p


def mark_size(p, size):
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)


def p_border(p, side="bottom", color=RULE, sz=6, space=4, val="single"):
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pBdr.append(el)


def tab_right(p, tw=CONTENT_TW):
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    t = OxmlElement("w:tab")
    t.set(qn("w:val"), "right")
    t.set(qn("w:pos"), str(tw))
    tabs.append(t)
    pPr.append(tabs)


def tab_stops(p, positions):
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    for pos in positions:
        t = OxmlElement("w:tab")
        t.set(qn("w:val"), "left")
        t.set(qn("w:pos"), str(int(pos * 20)))
        tabs.append(t)
    pPr.append(tabs)


def cell_props(cell, fill=None, borders=None, margins=None, width=None, valign=None):
    tcPr = cell._tc.get_or_add_tcPr()
    if width:
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is None:
            tcW = OxmlElement("w:tcW")
            tcPr.append(tcW)
        tcW.set(qn("w:w"), str(width))
        tcW.set(qn("w:type"), "dxa")
    for tag_name in ("w:tcBorders", "w:shd", "w:tcMar", "w:vAlign"):
        for old in tcPr.findall(qn(tag_name)):
            tcPr.remove(old)
    b = OxmlElement("w:tcBorders")
    spec = borders or {}
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge in spec:
            color, sz = spec[edge]
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:color"), color)
            el.set(qn("w:space"), "0")
        else:
            el.set(qn("w:val"), "nil")
        b.append(el)
    tcPr.append(b)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)
    if margins:
        m = OxmlElement("w:tcMar")
        for edge, v in zip(("top", "left", "bottom", "right"), margins):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:w"), str(v))
            el.set(qn("w:type"), "dxa")
            m.append(el)
        tcPr.append(m)
    if valign:
        va = OxmlElement("w:vAlign")
        va.set(qn("w:val"), valign)
        tcPr.append(va)


def make_table(doc, widths, rows=1):
    t = doc.add_table(rows=rows, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    tw = tblPr.find(qn("w:tblW"))
    if tw is None:
        tw = OxmlElement("w:tblW")
        tblPr.append(tw)
    tw.set(qn("w:w"), str(sum(widths)))
    tw.set(qn("w:type"), "dxa")
    # zero table-level cell margins & no indent
    mar = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    tblPr.append(ind)
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))
    for row in t.rows:
        for c, w in zip(row.cells, widths):
            cell_props(c, width=w)
    return t


def row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    trPr.append(el)


def row_height(row, pt):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(pt * 20)))
    h.set(qn("w:hRule"), "exact")
    trPr.append(h)


def first_par(cell):
    p = cell.paragraphs[0]
    return p


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
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
               "tblLook"]


def _reorder(el, order):
    def key(child):
        name = child.tag.split("}")[1]
        return order.index(name) if name in order else len(order)
    kids = list(el)
    for k in kids:
        el.remove(k)
    for k in sorted(kids, key=key):
        el.append(k)


def fix_order(root):
    for tag_name, order in (("w:pPr", PPR_ORDER), ("w:rPr", RPR_ORDER), ("w:tcPr", TCPR_ORDER), ("w:tblPr", TBLPR_ORDER)):
        for el in root.iter(qn(tag_name)):
            _reorder(el, order)


def spacer(doc, pt):
    p = doc.add_paragraph()
    fmt(p, 0, 0, line=pt)
    mark_size(p, 1)
    return p


def field(p, instr, size=8.5, color=MUTED):
    for kind, text in (("begin", None), (None, instr), ("separate", None), ("text", "1"), ("end", None)):
        r = p.add_run()
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.color.rgb = rgb(color)
        if kind in ("begin", "separate", "end"):
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            r._r.append(fc)
        elif kind is None:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = f" {text} "
            r._r.append(it)
        else:
            t = OxmlElement("w:t")
            t.text = text
            r._r.append(t)


# ------------------------------------------------------------------ photo
def make_photo():
    pdf = fitz.open(PDF)
    xref = pdf[0].get_images(full=True)[0][0]
    pix = fitz.Pixmap(pdf, xref)
    tmp = Path(tempfile.gettempdir())
    raw = tmp / "gem_photo_raw.png"
    pix.save(str(raw))
    im = Image.open(raw).convert("RGB")
    S = 600
    ring = 14
    im = im.resize((S, S), Image.LANCZOS)
    out = Image.new("RGBA", (S + 2 * ring, S + 2 * ring), (0, 0, 0, 0))
    big = Image.new("L", (S * 2, S * 2), 0)
    ImageDraw.Draw(big).ellipse((0, 0, S * 2 - 1, S * 2 - 1), fill=255)
    mask = big.resize((S, S), Image.LANCZOS)
    bigr = Image.new("L", ((S + 2 * ring) * 2,) * 2, 0)
    ImageDraw.Draw(bigr).ellipse((0, 0, (S + 2 * ring) * 2 - 1, (S + 2 * ring) * 2 - 1), fill=255)
    maskr = bigr.resize((S + 2 * ring, S + 2 * ring), Image.LANCZOS)
    out.paste(Image.new("RGBA", out.size, (0x25, 0x63, 0xEB, 255)), (0, 0), maskr)
    out.paste(im, (ring, ring), mask)
    path = tmp / "gem_photo_round.png"
    out.save(path)
    return path


# --------------------------------------------------------------- components
def section_heading(doc, text, before=12, page_break=False):
    p = doc.add_paragraph()
    fmt(p, before, 8, line=14, keep_next=True)
    if page_break:
        p.paragraph_format.page_break_before = True
    run(p, text.upper(), size=10.5, color=DARK, bold=True, spacing=0.8)
    p_border(p, "bottom", RULE, 6, 3)
    return p


def header(doc, photo):
    t = make_table(doc, [8300, 1580])
    left, right = t.rows[0].cells
    cell_props(left, valign="top")
    cell_props(right, valign="top")

    p = left.paragraphs[0]
    fmt(p, 0, 3, line=24)
    run(p, "SUMON RAY", size=20, color=DARK, bold=True)

    p = left.add_paragraph()
    fmt(p, 0, 7, line=14)
    run(p, "IT & DIGITALIZATION | ERP, INFRASTRUCTURE, SECURITY & AUTOMATION", size=11, color=BLUE, bold=True)

    def contact_line(items, stops):
        p = left.add_paragraph()
        fmt(p, 0, 4, line=11.5)
        tab_stops(p, stops)
        for i, (icon, text) in enumerate(items):
            if i:
                run(p, "\t", size=8.5)
            run(p, icon + " ", size=8.5, color=SLATE, font=EMOJI)
            run(p, text, size=8.5, color=INK)

    contact_line([("\U0001F4CD", "Dhaka, Bangladesh"), ("\u2709\ufe0f", "sumancsediu@gmail.com"), ("\U0001F4DE", "+880 1952 499680")], [100, 222])
    contact_line([("\U0001F517", "linkedin.com/in/sshanto13"), ("\U0001F310", "sshanto13.github.io"), ("\U0001F4BB", "github.com/sshanto13")], [126, 226])

    p = right.paragraphs[0]
    fmt(p, 6, 0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p.add_run().add_picture(str(photo), width=Inches(0.86))

    # dark rule under header
    p = doc.add_paragraph()
    fmt(p, 0, 0, line=8)
    mark_size(p, 2)
    p_border(p, "bottom", DARK, 16, 1)


def summary_box(doc, text):
    t = make_table(doc, [CONTENT_TW])
    c = t.rows[0].cells[0]
    cell_props(c, fill=BG, borders={"left": (BLUE, 16)}, margins=(140, 200, 140, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 0, line=12.7)
    run(p, text, size=9, color=BODY)


def small_cards_2x2(doc, cards):
    gap = 160
    w = (CONTENT_TW - gap) // 2
    t = make_table(doc, [w, gap, w], rows=3)
    spots = [(0, 0), (0, 2), (2, 0), (2, 2)]
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    for (r, c), (title, body) in zip(spots, cards):
        cell = t.rows[r].cells[c]
        cell_props(cell, borders=border, margins=(120, 130, 120, 130), width=w)
        p = cell.paragraphs[0]
        fmt(p, 0, 1, line=11)
        run(p, title, size=8.5, color=DARK, bold=True)
        p = cell.add_paragraph()
        fmt(p, 0, 0, line=10.8)
        run(p, body, size=8.5, color=INK)
    for r in (0, 2):
        row_cant_split(t.rows[r])
    for c in (0, 1, 2):
        mid = t.rows[1].cells[c]
        cell_props(mid, width=[w, gap, w][c])
        row_height(t.rows[1], 8)
        pp = mid.paragraphs[0]
        fmt(pp, 0, 0, line=8)
        mark_size(pp, 2)
    for c in (1,):
        for r in (0, 2):
            cell_props(t.rows[r].cells[c], width=gap)
            pp = t.rows[r].cells[c].paragraphs[0]
            fmt(pp, 0, 0, line=8)
            mark_size(pp, 2)


def profile_card(doc, title, lines):
    t = make_table(doc, [CONTENT_TW])
    row_cant_split(t.rows[0])
    c = t.rows[0].cells[0]
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    cell_props(c, borders=border, margins=(95, 160, 95, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 1, line=13)
    run(p, title, size=9.5, color=NAVY, bold=True)
    for label, text in lines:
        p = c.add_paragraph()
        fmt(p, 0, 0, line=12.5)
        run(p, label + ": ", size=8.5, color=DARK, bold=True)
        run(p, text, size=8.5, color=BODY)


def role_head(doc, title, dates, company, about, before=10):
    p = doc.add_paragraph()
    fmt(p, before, 0, line=13, keep_next=True)
    tab_right(p)
    run(p, title, size=10, color=DARK, bold=True)
    run(p, "\t" + dates, size=8.5, color=MUTED)
    p = doc.add_paragraph()
    fmt(p, 0, 1, line=12.5, keep_next=True)
    run(p, company.upper(), size=9.5, color=BLUE, bold=True)
    if about:
        p = doc.add_paragraph()
        fmt(p, 0, 2, line=11.5, keep_next=True)
        run(p, about, size=8.5, color=SLATE, italic=True)


def bullet(doc, label, text, keep_next=False):
    p = doc.add_paragraph(style="List Bullet")
    fmt(p, 0, 3.5, line=11, keep_next=keep_next, left=11, hanging=11)
    mark_size(p, 8.5)
    if label:
        run(p, label + " ", size=8.5, color=DARK, bold=True)
    run(p, text, size=8.5, color=BODY)
    return p


def tag(doc, text):
    p = doc.add_paragraph()
    fmt(p, 3, 0, line=12, left=4)
    run(p, "\u00A0Tech:\u00A0", size=8, color=INK, bold=True, shade=TAG)
    run(p, text + "\u00A0", size=8, color=SLATE, shade=TAG)
    return p


def plain_line(doc, text, size=8.5, color=BODY, italic=False, before=0, after=2, line=11.5, keep_next=False):
    p = doc.add_paragraph()
    fmt(p, before, after, line=line, keep_next=keep_next)
    run(p, text, size=size, color=color, italic=italic)
    return p


# -------------------------------------------------------------------- build
def build():
    photo = make_photo()
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(8.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    s = doc.sections[0]
    s.page_width, s.page_height = Emu(7560310), Emu(10692130)   # A4
    s.left_margin = s.right_margin = Pt(51)
    s.top_margin = Pt(TOP)
    s.bottom_margin = Pt(54)
    s.footer_distance = Pt(30)
    s.header_distance = Pt(20)

    # ------------------------------------------------------------ PAGE 1
    header(doc, photo)

    section_heading(doc, "Professional Summary", before=6)
    summary_box(
        doc,
        "IT and Digitalization leader with 12+ years of experience spanning enterprise application "
        "architecture, custom ERP development, IT infrastructure operations, cybersecurity, and digital "
        "transformation in Bangladesh, the US, and European markets. Combines hands-on technical expertise "
        "with business-focused delivery across FMCG, telecom, host infrastructure, and e-commerce sectors.",
    )
    spacer(doc, 8)
    small_cards_2x2(
        doc,
        [
            ("Global & MNC Leadership", "Multi-country leadership across US firms (Destinique, NybSys, Buzfi), European consultancies, and top BD enterprise groups."),
            ("Enterprise Systems & ERP", "End-to-end design, development, and scaling of custom ERP, POS, e-commerce, and supply chain management systems."),
            ("IT Infrastructure & Governance", "Strategy, policies, data center operations, network management, server hardening, and disaster recovery planning."),
            ("Cybersecurity & Risk Control", "Proactive vulnerability management, access controls, firewall configuration, backups, and comprehensive IT audits."),
        ],
    )

    section_heading(doc, "Professional Technical Profile", before=6)
    cards = [
        ("1. Enterprise Applications & Supply Chain Systems", [
            ("ERP & Operations", "ERP/POS systems, end-to-end inventory management, purchasing workflows, supply chain tracking, order fulfillment, multi-tier distributor/customer management."),
            ("Commercial Architecture", "Billing, invoicing, digital payments, booking systems, property/asset management workflows, manufacturing/distribution data systems."),
            ("Platforms & E-Commerce", "Shopify, WooCommerce, Wix, WordPress REST API, customized web-portal architecture."),
        ]),
        ("2. IT Infrastructure, Cloud & Cybersecurity", [
            ("Cloud Architecture", "AWS (EC2, RDS, SES, IAM), DigitalOcean, Vultr, Contabo VPS, scalable server configurations."),
            ("Core Infrastructure", "Linux server administration, DNS management, corporate email hosting, SSL/TLS protocols, firewall configuration, routing."),
            ("Security & Continuity", "Disaster recovery planning, automated backups and data recovery, access control management (MFA, RBAC), IT audits, risk assessment, production support."),
        ]),
        ("3. Software Engineering, Integration & Automation", [
            ("Backend & APIs", "PHP, Laravel, Node.js, Python, TypeScript, JavaScript, RESTful APIs, third-party web services, enterprise system integration."),
            ("Frontend Technologies", "React, Vue.js, Angular, Inertia.js, TypeScript, responsive UI frameworks (Bootstrap, Tailwind CSS)."),
            ("Workflow Automation", "Process automation tools, custom middleware development, agile project execution."),
        ]),
        ("4. Data, BI & Analytics", [
            ("Business Intelligence", "Power BI, operational dashboards, custom SQL reporting for executive decision-making."),
            ("Tracking & Performance", "Google Analytics 4 (GA4), data tracking pipelines, conversion metrics, performance analytics."),
            ("Database Management", "MySQL, PostgreSQL, advanced query optimization, data modeling, Redis caching."),
        ]),
        ("5. IT Governance, Digitalization & Leadership Tools", [
            ("Agile & AI Acceleration", "Utilizing AI-assisted tools (Cursor, Claude) for rapid MVP delivery, prototyping, and code reviews."),
        ]),
    ]
    for i, (title, lines) in enumerate(cards):
        profile_card(doc, title, lines)
        if i < len(cards) - 1:
            spacer(doc, 5)

    # ------------------------------------------------------------ PAGE 2
    section_heading(doc, "Professional Experience", before=0, page_break=True)

    role_head(doc, "Software Engineer", "Feb 2025 \u2013 Present | Florida, USA", "Destinique Travel LLC",
              "Florida-based luxury travel and vacation rental business providing digital services for travel and property operations.", before=2)
    bullet(doc, "", "Develop and maintain web applications and business automation solutions supporting travel inquiries, property quotations, and operational workflows.")
    bullet(doc, "", "Build frontend features, backend integrations, and API-driven functionality, supporting deployments, troubleshooting, and continuous enhancement.")
    bullet(doc, "Key Projects:", "Destinique Platform (destinique.com), Quote Application (quote.destinique.com), and Core Backend API Services.", keep_next=True)
    tag(doc, "Angular, PHP, Laravel, MySQL, JavaScript, REST APIs, Bootstrap, Cloudflare, Contabo VPS")

    role_head(doc, "Technical Manager, JS ERP (2023\u20132025) | Deputy Manager (2021\u20132023)", "Jan 2021 \u2013 Jan 2025 | Dhaka, Bangladesh",
              "JS Group \u2014 JS Enterprise / PaiMart",
              "Major FMCG importer & distributor with nationwide coverage, 1,000+ SKUs, 30,000+ general-trade outlets, and e-commerce operations.")
    bullet(doc, "", "Promoted from Deputy Manager; independently designed and developed JS ERP, unifying sales, purchasing, inventory, supplier/customer management, invoicing, and billing.")
    bullet(doc, "", "Managed complete SDLC from requirements gathering and process mapping through deployment, user support, and system stabilization.")
    bullet(doc, "", "Collaborated across sales, warehouse, and accounts teams to digitize workflows, enforce access control/MFA, and manage IT assets.")
    bullet(doc, "Ecosystem Portfolio:", "JS ERP, PaiPOS, Operation fulfillment system, Paimart e-commerce, Paibridge multivendor platform, and Paisoftdiaper.com.", keep_next=True)
    tag(doc, "PHP, Laravel, MySQL, JavaScript, REST APIs, Git, VPS Hosting, AWS, DigitalOcean, Vultr")

    role_head(doc, "Assistant Manager, Engineering", "Oct 2019 \u2013 Jan 2021 | Dhaka, Bangladesh", "ADN Telecom PLC",
              "Premier telecom, connectivity, and enterprise ICT solutions provider in Bangladesh.")
    bullet(doc, "", "Coordinated corporate client deployments from requirements gathering through integration, vendor alignment, and ongoing support.")
    bullet(doc, "", "Contributed to key in-house platforms (RoboKet, ADN SMS, ADN Email, ADN Server) and server configuration.")
    bullet(doc, "", "Delivered digital workflow and service management systems for government projects including NGTE Welfare Trust.")
    bullet(doc, "Key Corporate Clients:", "Nitol Niloy (NTRAC), Goods in Motion (GIM), Nadia Furniture, Esquire Electronics, FHDL.", keep_next=True)
    tag(doc, "PHP, Laravel, MySQL, JavaScript, Vultr VPS, DNS Management, Firewall Security")

    role_head(doc, "Software Engineer", "Dec 2017 \u2013 Sep 2019 | US / International", "NybSys Inc.",
              "USA-based technology and digital infrastructure provider operating across South Africa & Eswatini.")
    bullet(doc, "", "Developed web frontends and API integrations collaborating with distributed international engineering teams.")
    bullet(doc, "Key Deliverables:", "CDAS (Central Deductions Administration System for Eswatini government payroll), SLAS (Student Loan Application System portal), and Drinkwell web platform.", keep_next=True)
    tag(doc, "PHP, JavaScript, HTML5, CSS3, REST APIs, Git")

    # ------------------------------------------------------------ PAGE 3 (new section => own footer)
    doc.sections[0].footer.is_linked_to_previous = False
    s2 = doc.add_section(WD_SECTION.NEW_PAGE)

    role_head(doc, "Web Developer and Server Admin", "2014 \u2013 2017 | Dhaka, Bangladesh", "Dhaka Web Host Ltd.",
              "Web hosting, domain infrastructure, and server solutions provider.", before=0)
    bullet(doc, "", "Administered Linux server environments supporting corporate clients, including DNS, email hosting, SSL certificates, and firewall hardening.")
    bullet(doc, "", "Maintained 99.9% uptime for enterprise clients including US-Bangla Airlines, Aftab Group, Bangla Academy, and Dhaka Life Insurance.", keep_next=True)
    tag(doc, "Linux, DNS Management, Corporate Email, SSL/TLS, Firewall Configuration, Hosting Admin")

    section_heading(doc, "Global & Local Consultancy Engagements", before=14)
    spacer(doc, 4)
    ent = [
        ("BUZFI LLC (Jamaica, NY, USA) | Head of Technology", "(Aug 2023 \u2013 Dec 2025)", "Led tech architecture for US digital marketplace startup: Next.js frontend on Vercel, Laravel API backend, Cloudflare, and custom VPS infrastructure."),
        ("Motmans Digital Vision (Belgium) | Head of Development", "(2020 \u2013 2023)", "Managed product engineering and business application delivery for Belgian corporate clients."),
        ("Rivr (Belgium) | Engineering Manager", "(2019 \u2013 2023)", "Directed web engineering for digital branding agency, including solutions for Orka Group (SAP services firm)."),
        ("Stck.Pro (Spain) | Technical Project Manager", "(2022 \u2013 2023)", "Managed technical roadmap for stock market news platform targeting US and EU markets."),
        ("Spiretec Limited & Sysba Trade Brand (Bangladesh) | Technical Consultant", "", "Provided direct architecture consultancy, digital infrastructure setup, and application development support."),
    ]
    t = make_table(doc, [CONTENT_TW])
    c = t.rows[0].cells[0]
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    cell_props(c, fill=BG, borders=border, margins=(140, 160, 120, 160), width=CONTENT_TW)
    first = True
    for i, (head, date, desc) in enumerate(ent):
        p = c.paragraphs[0] if first else c.add_paragraph()
        first = False
        fmt(p, 0 if i == 0 else 5, 1, line=11.5, keep_next=True)
        run(p, head + (" " if date else ""), size=8.5, color=DARK, bold=True)
        if date:
            run(p, date, size=8, color=MUTED)
        p = c.add_paragraph()
        fmt(p, 0, 4 if i < len(ent) - 1 else 0, line=11.3)
        run(p, desc, size=8.5, color=INK)
        if i < len(ent) - 1:
            p_border(p, "bottom", RULE, 4, 5, val="dotted")

    section_heading(doc, "Government Initiatives & Training", before=14)
    p = doc.add_paragraph()
    fmt(p, 0, 0, line=13, keep_next=True)
    tab_right(p)
    run(p, "Trainer \u2014 ICT Division (Learning & Earning Development Project)", size=10, color=DARK, bold=True)
    run(p, "\tDec 2016 \u2013 Apr 2017", size=8.5, color=MUTED)
    plain_line(doc, "Government initiative focused on ICT skills development and capacity building.", color=SLATE, italic=True, after=2, keep_next=True)
    plain_line(doc, "Trained over 200 software developers in HTML, CSS, PHP, WordPress, and MySQL through practical lab sessions.")

    section_heading(doc, "Education & Professional Certifications", before=14)
    spacer(doc, 4)
    # degree card
    t = make_table(doc, [CONTENT_TW])
    c = t.rows[0].cells[0]
    cell_props(c, fill=BG, borders=border, margins=(120, 160, 120, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 1, line=13)
    tab_right(p, CONTENT_TW - 320)
    run(p, "Bachelor of Science (B.Sc.) in Computer Science & Engineering", size=9.5, color=DARK, bold=True)
    run(p, "\t2010 \u2013 2014", size=8.5, color=MUTED)
    p = c.add_paragraph()
    fmt(p, 0, 0, line=11.5)
    run(p, "Dhaka International University", size=8.5, color=BLUE, bold=True)
    spacer(doc, 6)

    gap = 160
    w = (CONTENT_TW - gap) // 2
    t = make_table(doc, [w, gap, w])
    for ci, (title, sub) in ((0, ("Web Application Development with PHP", "BASIS (Bangladesh Association of Software & Information Services) \u2014 2015")),
                             (2, ("Web Design & Development", "Daffodil Institute of Information Technology (DIIT) \u2014 2012"))):
        cell = t.rows[0].cells[ci]
        cell_props(cell, borders=border, margins=(110, 130, 110, 130), width=w)
        p = cell.paragraphs[0]
        fmt(p, 0, 1, line=11)
        run(p, title, size=8.5, color=DARK, bold=True)
        p = cell.add_paragraph()
        fmt(p, 0, 0, line=10)
        run(p, sub, size=8, color=MUTED)
    pp = t.rows[0].cells[1].paragraphs[0]
    fmt(pp, 0, 0, line=8)
    mark_size(pp, 2)
    spacer(doc, 6)

    t = make_table(doc, [CONTENT_TW])
    c = t.rows[0].cells[0]
    cell_props(c, fill=TAG, margins=(100, 160, 100, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 0, line=11.5)
    run(p, "Languages: ", size=8.5, color=DARK, bold=True)
    run(p, "Bangla (Native) | English (Professional Working Proficiency)", size=8.5, color=INK)

    # ------------------------------------------------------------ footers
    def footer_par(footer, left_text=None):
        p = footer.paragraphs[0]
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        fmt(p, 0, 0, line=12)
        tab_right(p)
        tabs = p._p.pPr.find(qn("w:tabs"))
        for pos in (4680, 9360):          # clear the Footer style's default centre/right tabs
            t = OxmlElement("w:tab")
            t.set(qn("w:val"), "clear")
            t.set(qn("w:pos"), str(pos))
            tabs.append(t)
        p_border(p, "top", CARD, 6, 6)
        if left_text:
            run(p, left_text, size=8.5, color=MUTED)
        run(p, "\tPage ", size=8.5, color=MUTED)
        field(p, "PAGE")
        run(p, " of ", size=8.5, color=MUTED)
        field(p, "NUMPAGES")

    footer_par(doc.sections[0].footer)
    s2.footer.is_linked_to_previous = False
    footer_par(s2.footer, "Sumon Ray \u2014 Executive Curriculum Vitae")
    s2.left_margin = s2.right_margin = Pt(51)
    s2.top_margin = Pt(TOP)
    s2.bottom_margin = Pt(54)
    s2.footer_distance = Pt(30)

    fix_order(doc.element)
    for sec in doc.sections:
        fix_order(sec.footer._element)

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_v2.docx")
        doc.save(alt)
        print("locked; saved", alt)


if __name__ == "__main__":
    build()

