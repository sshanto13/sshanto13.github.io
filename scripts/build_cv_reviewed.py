"""Rebuild the CV from the (approved) content of public/files/text.docx using the existing
Gemini-style layout, with a refined Arial + Calibri typographic system.
Content is copied from text.docx; only typography/layout is decided here."""
import copy
import os
import sys
import tempfile
from pathlib import Path

import fitz
import win32com.client as w
from PIL import Image, ImageDraw
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
PHOTO_SRC_PDF = BASE / "Sumon Ray- ACI.pdf"
OUT = BASE / "Sumon_Ray_CV_Reviewed.docx"
TMP = Path(tempfile.gettempdir())

HEADF = "Arial"
BODYF = "Calibri"
EMOJI = "Segoe UI Emoji"

DARK = "0F172A"
INK = "1E293B"
BODY = "334155"
SLATE = "475569"
MUTED = "64748B"
BLUE = "2563EB"
NAVY = "14305A"
RULE = "CBD5E1"
CARD = "E2E8F0"
BG = "F8FAFC"
TAG = "F1F5F9"

CONTENT_TW = 9880          # 494 pt
BOTTOM_LIMIT = 776

# global spacing scale used by helpers (set by build())
G = 1.0
K = 1.0

# ----------------------------------------------------------------------------- content
SUMMARY = (
    "IT and Digitalization professional with over a decade of experience in software engineering, business "
    "applications, custom ERP development, e-commerce platforms, system integration, and IT infrastructure "
    "support. Combines hands-on development experience with business process automation, operational "
    "reporting, hosting management, and cross-functional project coordination. Experienced in delivering "
    "practical technology solutions for local and global clients across e-commerce, distribution, "
    "telecommunications, travel, and property management."
)

HIGHLIGHTS = [
    ("Local & Global Delivery", "Technology solutions delivered for employers and clients in Bangladesh, the US, and Europe across e-commerce, distribution, telecom, and travel."),
    ("Enterprise Systems & ERP", "Design, development, and support of custom ERP, POS, e-commerce, and supply chain management systems."),
    ("Hosting & Infrastructure", "Server administration, DNS, email hosting, backups, access control, and hosting cost management."),
    ("Security Operations", "Firewall configuration, Cloudflare-based Layer 7 DDoS mitigation, access controls, and automated backups."),
]

PROFILE = [
    ("1. Enterprise Applications & Supply Chain Systems", [
        ("ERP & Operations", "ERP/POS systems, end-to-end inventory management, purchasing workflows, supply chain, tracking, order fulfillment, multi-tier distributor/customer management."),
        ("Commercial Architecture", "Billing, invoicing, digital payments, booking systems, property/asset management workflows, manufacturing/distribution data systems."),
        ("Platforms & E-Commerce", "Shopify, WooCommerce, Wix, WordPress REST API, customized web-portal architecture."),
    ]),
    ("2. IT Infrastructure, Cloud & Cybersecurity", [
        ("Cloud Architecture", "AWS (EC2, RDS, SES, IAM), DigitalOcean, Vultr, Contabo VPS, scalable server configurations."),
        ("Core Infrastructure", "Linux server administration, DNS management, corporate email hosting, SSL/TLS protocols, firewall configuration, routing."),
        ("Security & Continuity", "Automated backups and data recovery, access control management (MFA, RBAC), firewall configuration, production support."),
    ]),
    ("3. Software Engineering, Integration & Automation", [
        ("Backend & APIs", "PHP, Laravel, Node.js, Python, TypeScript, JavaScript, RESTful APIs, third-party web services, enterprise system integration."),
        ("Frontend Technologies", "React, Vue.js, Angular, Inertia.js, TypeScript, responsive UI frameworks (Bootstrap, Tailwind CSS)."),
        ("Workflow Automation", "Process automation tools, custom middleware development, agile project execution."),
        ("Agile & AI Acceleration", "Utilizing AI-assisted tools (Cursor, Claude) for rapid MVP delivery, prototyping, and code reviews."),
    ]),
    ("4. Data, BI & Analytics", [
        ("Business Intelligence", "Power BI, operational dashboards, custom SQL reporting for executive decision-making."),
        ("Tracking & Performance", "Google Analytics 4 (GA4), data tracking pipelines, conversion metrics, performance analytics."),
        ("Database Management", "MySQL, PostgreSQL, advanced query optimization, data modeling, Redis caching."),
    ]),
    ("5. Leadership & Delivery", [
        ("Team & Delivery Coordination", "Experienced in coordinating frontend, backend, and QA teams, aligning technical execution with business requirements, and managing project priorities across development and delivery stages. Skilled in collaborating with business stakeholders, facilitating cross-functional communication, and resolving technical dependencies."),
        ("Vendor & Technology Coordination", "Experienced in coordinating with technology vendors and service providers for server, cloud infrastructure, and technical services. Skilled in evaluating technology solutions, assessing infrastructure requirements, troubleshooting vendor-related issues, and balancing cost, performance, reliability, and business needs."),
    ]),
]

EXPERIENCE = [
    dict(
        title="Software Engineer \u2014 Business Applications & Automation", company="DESTINIQUE TRAVEL LLC",
        date="Feb 2025 \u2013 Present | Remote",
        bullets=[
            "Develop and maintain web applications, REST API integrations, and business automation solutions for travel inquiries, property quotations, visitor tracking, and operational workflows.",
            "Built custom tools for ad generation, promotion management, and monitoring, reducing repetitive marketing tasks and improving workflow efficiency.",
            "Automated property-description generation for quotation workflows, reducing manual writing and supporting the quotation team.",
            "Develop frontend features, backend integrations, and API-driven functionality while supporting deployment, troubleshooting, and continuous improvement.",
            "Monitor hosting costs and evaluate infrastructure options to balance performance, reliability, and operating expenses.",
        ],
        extras=[("Key Projects", "Destinique Platform, Quote Application, Ad Generation Tool, Promotion Management Tool, Monitoring Tool, and Core Backend API Services.")],
        tech="Angular, PHP, Laravel, MySQL, JavaScript, REST APIs, Bootstrap, Tailwind CSS, Cloudflare, Contabo VPS.",
    ),
    dict(
        title="Technical Manager \u2014 JS ERP", company="JS GROUP", date="Jan 2023 \u2013 Jan 2025",
        bullets=[
            "Independently designed and developed JS ERP, replacing the previous reliance on Tally for key business operations and centralizing sales, purchasing, inventory, supplier and customer management, invoicing, and billing.",
            "Automated key business processes through ERP workflows, improving operational coordination and visibility across commercial functions.",
            "Developed sales-force and distribution workflows to track distributor sales and monitor sales-officer performance, supporting field operations and management reporting.",
            "Managed the ERP development lifecycle, including requirements gathering, business-process mapping, database design, development, deployment, user support, and system stabilization.",
            "Coordinated with sales, warehouse, accounts, and technical stakeholders to translate business requirements into practical systems and streamline workflows.",
            "Managed the development and maintenance of ERP, POS, order-fulfilment, and internal business applications supporting retail, wholesale, inventory, and commercial operations.",
        ],
    ),
    dict(
        title="Deputy Manager \u2014 PaiMart", company="JS GROUP", date="Jan 2021 \u2013 Jan 2023",
        bullets=[
            "Supported the development and operation of PaiMart's e-commerce platform, serving more than 200,000 users, including application functionality, integrations, and ongoing technical improvements.",
            "Contributed to platform reliability, website traffic management, and security, including Cloudflare configuration to help mitigate Layer 7 application-layer attacks.",
            "Managed server billing and evaluated hosting options, including Vultr and AWS, to balance infrastructure costs, performance, and reliability.",
            "Collaborated with business and technical stakeholders to support e-commerce operations, order fulfilment, inventory workflows, and related business applications.",
        ],
        extras=[("Key Projects", "JS ERP, PaiPOS, PaiMart E-commerce (200k+ users), Operations Fulfilment System, PaiBridge Multivendor Platform, Paisoftdiaper.com.")],
        tech="PHP, Laravel, MySQL, JavaScript, REST APIs, Git, Cloudflare, Linux/VPS Hosting, AWS, DigitalOcean, Vultr.",
    ),
    dict(
        title="Assistant Manager \u2014 Engineering", company="ADN TELECOM PLC", date="Oct 2019 \u2013 Jan 2021 | Dhaka, Bangladesh",
        bullets=[
            "Coordinated corporate technology deployments from requirements gathering and system integration to vendor coordination and ongoing technical support.",
            "Collaborated with cross-functional teams to analyse client requirements and deliver production-ready applications and digital services.",
            "Contributed to technology products and services including RoboKet, ADN SMS, ADN Email, and ADN Server, with involvement in server configuration and support.",
            "Delivered digital workflow and service-management solutions for clients, including the Non-Government Teachers and Employees Welfare Trust (NGTE Welfare Trust), a Bangladesh government ICT initiative.",
            "Supported client projects for organizations including Nitol Niloy (NTRAC), Goods in Motion (GIM), Nadia Furniture, Esquire Electronics, and FHDL.",
        ],
        tech="PHP, Laravel, MySQL, JavaScript, VPS Hosting, DNS, Firewall Configuration.",
    ),
    dict(
        title="Software Engineer", company="NYBSYS INC.", date="Dec 2017 \u2013 Sep 2019 | Dhaka, Bangladesh",
        bullets=[
            "Developed web applications, responsive interfaces, and API integrations for business and public-service solutions.",
            "Collaborated with distributed teams on projects involving web, mobile, IoT, LTE, and telecommunications-related solutions.",
            "Translated UI/UX designs and business requirements into functional, production-ready web applications.",
            "Worked with backend and engineering teams to integrate APIs and support end-to-end product functionality.",
            "Contributed to CDAS (Central Deductions Administration System for government payroll deductions in Eswatini), SLAS (Student Loan Application System portal), and the Drinkwell web platform.",
        ],
        tech="PHP, Laravel, JavaScript, HTML5, CSS3, REST APIs, Bootstrap, Git.",
    ),
    dict(
        title="Web Developer & Server Administrator", company="DHAKA WEB HOST LTD.", date="Dec 2014 \u2013 Nov 2017 | Dhaka, Bangladesh",
        bullets=[
            "Contributed to building and developing the web development team, supporting project delivery, client requirements, and the growth of the company's web development services.",
            "Developed and maintained client websites and web applications while providing technical support for Linux hosting environments.",
            "Administered hosting services, server configurations, DNS, SSL/TLS, corporate email, firewalls, and troubleshooting to support reliable website operations.",
            "Combined web development and server administration responsibilities to deliver end-to-end solutions for client websites and hosting services.",
            "Supported projects and services for clients including US-Bangla Airlines, Aftab Group, Bangla Academy, and Dhaka Life Insurance.",
        ],
        tech="Web Development, Linux, DNS, SSL/TLS, Corporate Email, Firewall Configuration, WHMCS, cPanel, Hosting Administration.",
    ),
]

CONSULTANCY = [
    dict(
        title="Head of Technology", company="BUZFI LLC", date="Consultancy | Aug 2023 \u2013 Dec 2025 | Jamaica, New York, USA",
        bullets=[
            "Contributed to technology delivery for a B2C and B2B e-commerce marketplace, including its Next.js frontend, Laravel backend APIs, and deployment infrastructure.",
            "Supported platform integrations, hosting, performance, and security across Vercel, Cloudflare, and VPS environments.",
            "Worked on technical improvements and platform maintenance to support reliable e-commerce operations.",
        ],
        tech="Next.js, Laravel, REST APIs, Cloudflare, Vercel, VPS.",
    ),
    dict(
        title="Head of Development", company="MOTMANS DIGITAL VISION / MOTMANS DIGITAL SOLUTIONS", date="Consultancy | 2020 \u2013 2023 | Belgium",
        bullets=[
            "Successfully delivered more than 50 web projects for business clients, covering website development, business applications, and digital solutions.",
            "Led and coordinated development activities to align client requirements, technical implementation, project priorities, and delivery timelines.",
            "Worked with business stakeholders to translate client requirements into practical digital solutions and support successful project delivery.",
            "Coordinated development resources and technical execution across multiple projects to maintain delivery quality and meet business expectations.",
        ],
    ),
    dict(
        title="Technical Project Manager", company="STCK.PRO", date="Consultancy | 2022 \u2013 2023 | Spain",
        bullets=[
            "Coordinated the delivery and publication of Stck.Pro applications for both iOS and Android, extending the platform's availability across mobile devices.",
            "Led and coordinated frontend, backend, and QA activities, aligning technical work with business requirements and delivery priorities.",
            "Worked across technical and business teams to coordinate project execution, resolve delivery dependencies, and support release readiness.",
            "Contributed to the delivery of a stock-market news platform targeting audiences in the US and European markets.",
        ],
    ),
    dict(
        title="Digital Business & Technology Consultant", company="SPIRETEC LIMITED / SYSBA TRADE BRAND", date="Consultancy | Bangladesh",
        bullets=[
            "Supported a business launch by coordinating brand identity, website development, mobile app initiatives, and digital marketing activities.",
            "Coordinated design, development, SEO, and promotional activities to establish the business's digital presence and support market entry.",
        ],
    ),
]


# ----------------------------------------------------------------------------- helpers
def rgb(h):
    return RGBColor.from_string(h)


def run(p, text, size=10, color=BODY, bold=False, italic=False, font=BODYF, shade=None, spacing=None):
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


def add_hyperlink(p, url, text, size=9.5, color=INK, font=BODYF):
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), font)
    rPr.append(rf)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rPr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    hl.append(r)
    p._p.append(hl)


def fmt(p, before=0, after=0, line=None, align=None, keep_next=False, left=None, hanging=None, scale=True):
    pf = p.paragraph_format
    g = G if scale else 1.0
    k = K if scale else 1.0
    pf.space_before = Pt(before * g)
    pf.space_after = Pt(after * g)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line * k)
    if align is not None:
        p.alignment = align
    pf.keep_with_next = keep_next
    pf.keep_together = True
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
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        tabs = OxmlElement("w:tabs")
        pPr.append(tabs)
    t = OxmlElement("w:tab")
    t.set(qn("w:val"), "right")
    t.set(qn("w:pos"), str(tw))
    tabs.append(t)


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
    for name in ("w:tcBorders", "w:shd", "w:tcMar", "w:vAlign"):
        for old in tcPr.findall(qn(name)):
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
            el.set(qn("w:w"), str(int(v)))
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
    for gc, wdt in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(wdt))
    for row in t.rows:
        for c, wdt in zip(row.cells, widths):
            cell_props(c, width=wdt)
    return t


def row_cant_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def row_height(row, pt):
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(pt * 20)))
    h.set(qn("w:hRule"), "exact")
    row._tr.get_or_add_trPr().append(h)


def spacer(doc, pt, keep_next=False):
    p = doc.add_paragraph()
    fmt(p, 0, 0, line=pt, keep_next=keep_next)
    mark_size(p, 1)
    return p


def field(p, instr, size=9, color=MUTED):
    for kind, text in (("begin", None), (None, instr), ("separate", None), ("text", "1"), ("end", None)):
        r = p.add_run()
        r.font.name = BODYF
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
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]


def _reorder(el, order):
    def key(c):
        n = c.tag.split("}")[1]
        return order.index(n) if n in order else len(order)
    kids = list(el)
    for k in kids:
        el.remove(k)
    for k in sorted(kids, key=key):
        el.append(k)


def fix_order(root):
    for tag, order in (("w:pPr", PPR_ORDER), ("w:rPr", RPR_ORDER), ("w:tcPr", TCPR_ORDER), ("w:tblPr", TBLPR_ORDER)):
        for el in root.iter(qn(tag)):
            _reorder(el, order)


def make_photo():
    pdf = fitz.open(PHOTO_SRC_PDF)
    xref = pdf[0].get_images(full=True)[0][0]
    pix = fitz.Pixmap(pdf, xref)
    raw = TMP / "cv_photo_raw.png"
    pix.save(str(raw))
    im = Image.open(raw).convert("RGB")
    S, ring = 600, 14
    im = im.resize((S, S), Image.LANCZOS)
    out = Image.new("RGBA", (S + 2 * ring, S + 2 * ring), (0, 0, 0, 0))
    big = Image.new("L", (S * 2, S * 2), 0)
    ImageDraw.Draw(big).ellipse((0, 0, S * 2 - 1, S * 2 - 1), fill=255)
    mask = big.resize((S, S), Image.LANCZOS)
    n = (S + 2 * ring) * 2
    bigr = Image.new("L", (n, n), 0)
    ImageDraw.Draw(bigr).ellipse((0, 0, n - 1, n - 1), fill=255)
    maskr = bigr.resize((S + 2 * ring, S + 2 * ring), Image.LANCZOS)
    out.paste(Image.new("RGBA", out.size, (0x25, 0x63, 0xEB, 255)), (0, 0), maskr)
    out.paste(im, (ring, ring), mask)
    path = TMP / "cv_photo_round.png"
    out.save(path)
    return path


# ----------------------------------------------------------------------------- components
def section_heading(doc, text, before=14):
    p = doc.add_paragraph()
    fmt(p, before, 7, line=15, keep_next=True)
    run(p, text.upper(), size=11, color=DARK, bold=True, font=HEADF, spacing=0.9)
    p_border(p, "bottom", RULE, 6, 3)
    return p


def header(doc, photo):
    t = make_table(doc, [8300, 1580])
    left, right = t.rows[0].cells
    cell_props(left, valign="top", width=8300)
    cell_props(right, valign="top", width=1580)

    p = left.paragraphs[0]
    fmt(p, 0, 3, line=31, scale=False)
    run(p, "SUMON RAY", size=26, color=DARK, bold=True, font=HEADF, spacing=0.4)

    p = left.add_paragraph()
    fmt(p, 0, 8, line=14, scale=False)
    run(p, "IT & DIGITALIZATION | ERP, INFRASTRUCTURE, SECURITY & AUTOMATION", size=10, color=BLUE, bold=True, font=HEADF)

    def contact_line(items, stops):
        p = left.add_paragraph()
        fmt(p, 0, 3, line=12.5, scale=False)
        tab_stops(p, stops)
        for i, (icon, kind, text, url) in enumerate(items):
            if i:
                run(p, "\t", size=9.5)
            run(p, icon + " ", size=9, color=SLATE, font=EMOJI)
            if url:
                add_hyperlink(p, url, text, size=9.5, color=INK)
            else:
                run(p, text, size=9.5, color=INK)

    contact_line([
        ("\U0001F4CD", "t", "Dhaka, Bangladesh", None),
        ("\u2709\ufe0f", "t", "sumancsediu@gmail.com", "mailto:sumancsediu@gmail.com"),
        ("\U0001F4DE", "t", "+880 1952 499680", "tel:+8801952499680"),
    ], [118, 250])
    contact_line([
        ("\U0001F310", "t", "sshanto13.github.io", "https://sshanto13.github.io"),
        ("\U0001F4BB", "t", "github.com/sshanto13", "https://github.com/sshanto13"),
    ], [118])

    p = right.paragraphs[0]
    fmt(p, 4, 0, align=WD_ALIGN_PARAGRAPH.RIGHT, scale=False)
    p.add_run().add_picture(str(photo), width=Inches(0.86))

    p = doc.add_paragraph()
    fmt(p, 0, 0, line=8, scale=False)
    mark_size(p, 2)
    p_border(p, "bottom", DARK, 16, 1)


def summary_box(doc, text):
    t = make_table(doc, [CONTENT_TW])
    row_cant_split(t.rows[0])
    c = t.rows[0].cells[0]
    cell_props(c, fill=BG, borders={"left": (BLUE, 16)}, margins=(130 * G, 200, 130 * G, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 0, line=14)
    run(p, text, size=10, color=BODY)


def highlight_cards(doc, cards):
    gap = 160
    wd = (CONTENT_TW - gap) // 2
    t = make_table(doc, [wd, gap, wd], rows=3)
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    for (r, c), (title, body) in zip([(0, 0), (0, 2), (2, 0), (2, 2)], cards):
        cell = t.rows[r].cells[c]
        cell_props(cell, borders=border, margins=(110 * G, 130, 110 * G, 130), width=wd)
        p = cell.paragraphs[0]
        fmt(p, 0, 1.5, line=12)
        run(p, title, size=9.5, color=NAVY, bold=True, font=HEADF)
        p = cell.add_paragraph()
        fmt(p, 0, 0, line=12.2)
        run(p, body, size=10, color=INK)
    for r in (0, 2):
        row_cant_split(t.rows[r])
    for r in range(3):
        for c in (0, 1, 2):
            cell = t.rows[r].cells[c]
            if c == 1 or r == 1:
                cell_props(cell, width=[wd, gap, wd][c])
                pp = cell.paragraphs[0]
                fmt(pp, 0, 0, line=8, scale=False)
                mark_size(pp, 2)
    row_height(t.rows[1], 8 * G)


def profile_card(doc, title, lines):
    t = make_table(doc, [CONTENT_TW])
    row_cant_split(t.rows[0])
    c = t.rows[0].cells[0]
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    cell_props(c, borders=border, margins=(100 * G, 160, 100 * G, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 2, line=14)
    run(p, title, size=10, color=NAVY, bold=True, font=HEADF)
    for label, text in lines:
        p = c.add_paragraph()
        fmt(p, 0, 1, line=13)
        if label:
            run(p, label + ": ", size=10, color=DARK, bold=True)
        run(p, text, size=10, color=BODY)


INTRO = {
    "DESTINIQUE TRAVEL LLC": "Florida, USA\u2013based travel and property management company.",
    "JS GROUP": "Bangladesh-based consumer goods import, distribution, and e-commerce.",
    "ADN TELECOM PLC": "Bangladesh-based telecom, connectivity, and enterprise ICT solutions.",
    "NYBSYS INC.": "USA-based technology, telecommunications, and digital infrastructure solutions.",
    "DHAKA WEB HOST LTD.": "Bangladesh-based web hosting, domains, and server solutions.",
    "BUZFI LLC": "Jamaica, New York, USA\u2013based startup digital marketplace connecting buyers and sellers.",
    "MOTMANS DIGITAL VISION / MOTMANS DIGITAL SOLUTIONS": "Belgium-based IT consultancy specializing in business applications and automation.",
}


def role_block(doc, e, first_in_section=False):
    parts = []   # collect paragraphs to set keep_with_next chain
    p = doc.add_paragraph()
    fmt(p, 3 if first_in_section else 11, 0, line=14, keep_next=True)
    tab_right(p)
    run(p, e["title"], size=10.5, color=DARK, bold=True, font=HEADF)
    run(p, "\t" + e["date"], size=9.5, color=MUTED)
    parts.append(p)
    p = doc.add_paragraph()
    intro = INTRO.get(e["company"])
    if e["title"].startswith("Deputy Manager"):
        intro = None        # same JS Group intro already shown on the previous entry
    fmt(p, 0, 1 if intro else 3, line=13, keep_next=True)
    run(p, e["company"], size=9.5, color=BLUE, bold=True, font=HEADF, spacing=0.3)
    parts.append(p)
    if intro:
        p = doc.add_paragraph()
        fmt(p, 0, 3, line=12.6, keep_next=True)
        run(p, intro, size=9.5, color=SLATE, italic=True)
        parts.append(p)
    for b in e.get("bullets", []):
        p = doc.add_paragraph(style="List Bullet")
        fmt(p, 0, 2.5, line=12.8, keep_next=True, left=11, hanging=11)
        mark_size(p, 10)
        run(p, b, size=10, color=BODY)
        parts.append(p)
    for label, text in e.get("extras", []):
        p = doc.add_paragraph(style="List Bullet")
        fmt(p, 0, 2.5, line=12.8, keep_next=True, left=11, hanging=11)
        mark_size(p, 10)
        run(p, label + ": ", size=10, color=DARK, bold=True)
        run(p, text, size=10, color=BODY)
        parts.append(p)
    if e.get("tech"):
        p = doc.add_paragraph()
        fmt(p, 3, 0, line=13, left=4)
        run(p, "\u00A0Technologies:\u00A0", size=9, color=INK, bold=True, shade=TAG)
        run(p, e["tech"] + "\u00A0", size=9, color=SLATE, shade=TAG)
        parts.append(p)
    # keep heading + company + first bullet + second bullet together and the last two items together;
    # the middle of a long entry may flow to the next page (never a lone line or lone bullet)
    n = len(parts)
    head_n = 4 if intro else 3
    for i, para in enumerate(parts):
        para.paragraph_format.keep_with_next = (i < head_n or i == n - 2) and i < n - 1
    return parts


def plain(doc, text, size=10, color=BODY, italic=False, before=0, after=2, line=13, keep_next=False):
    p = doc.add_paragraph()
    fmt(p, before, after, line=line, keep_next=keep_next)
    run(p, text, size=size, color=color, italic=italic)
    return p


def footer_setup(section, total_pages_field=True):
    f = section.footer
    p = f.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    fmt(p, 0, 0, line=12, scale=False)
    tab_right(p)
    pPr = p._p.get_or_add_pPr()
    tabs = pPr.find(qn("w:tabs"))
    for pos in (4680, 9360):
        t = OxmlElement("w:tab")
        t.set(qn("w:val"), "clear")
        t.set(qn("w:pos"), str(pos))
        tabs.append(t)
    p_border(p, "top", CARD, 6, 6)
    run(p, "Sumon Ray", size=9, color=MUTED)
    run(p, "\tPage ", size=9, color=MUTED)
    field(p, "PAGE")
    run(p, " of ", size=9, color=MUTED)
    field(p, "NUMPAGES")


# ----------------------------------------------------------------------------- build
def join_last_two_words(para):
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


def rescale_body(doc, line_f):
    """Body 10 -> 9.5, small 9.5 -> 9, tags 9 -> 8.5 (header block untouched)."""
    mapping = {20: 19, 19: 18, 18: 17}          # half-points
    first_tbl = doc.tables[0]._tbl
    for p in doc.element.body.iter(qn("w:p")):
        if first_tbl in list(p.iterancestors()):
            continue
        changed = False
        for sz in p.iter(qn("w:sz")):
            v = int(sz.get(qn("w:val")))
            if v in mapping:
                sz.set(qn("w:val"), str(mapping[v]))
                changed = True
        if changed:
            sp = p.find(qn("w:pPr") + "/" + qn("w:spacing"))
            if sp is not None and sp.get(qn("w:lineRule")) == "exact":
                sp.set(qn("w:line"), str(int(int(sp.get(qn("w:line"))) * line_f)))


def build(g=1.0, k=1.0, nbsp=(), page_breaks=(), small=False):
    global G, K
    G, K = g, k
    photo = make_photo()
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = BODYF
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODYF)

    s = doc.sections[0]
    s.page_width, s.page_height = Emu(7560310), Emu(10692130)  # A4
    s.left_margin = s.right_margin = Pt(51)
    s.top_margin = Pt(50)
    s.bottom_margin = Pt(54)
    s.footer_distance = Pt(28)

    header(doc, photo)

    def heading(text, before=14, key=None):
        p = section_heading(doc, text, before)
        if key and key in page_breaks:
            p.paragraph_format.page_break_before = True
        return p

    heading("Professional Summary", before=9)
    spacer(doc, 4, keep_next=True)
    summary_box(doc, SUMMARY)
    spacer(doc, 8)
    highlight_cards(doc, HIGHLIGHTS)

    heading("Professional Technical Profile", before=10, key="profile")
    spacer(doc, 4, keep_next=True)
    for i, (title, lines) in enumerate(PROFILE):
        profile_card(doc, title, lines)
        if i < len(PROFILE) - 1:
            spacer(doc, 5 * g)

    heading("Professional Experience", before=14, key="experience")
    spacer(doc, 2, keep_next=True)
    for i, e in enumerate(EXPERIENCE):
        role_block(doc, e, first_in_section=(i == 0))

    heading("Selected Consultancy & Project Engagements", before=16, key="consultancy")
    spacer(doc, 2, keep_next=True)
    for i, e in enumerate(CONSULTANCY):
        role_block(doc, e, first_in_section=(i == 0))

    heading("Government Initiatives & Training", before=16, key="government")
    p = doc.add_paragraph()
    fmt(p, 3, 0, line=14, keep_next=True)
    tab_right(p)
    run(p, "Trainer \u2014 ICT Division (Learning & Earning Development Project)", size=10.5, color=DARK, bold=True, font=HEADF)
    run(p, "\tDec 2016 \u2013 Apr 2017 | Dhaka, Bangladesh", size=9.5, color=MUTED)
    plain(doc, "Government initiative focused on ICT skills development and capacity building.", color=SLATE, italic=True, after=2.5, keep_next=True, before=1)
    b = doc.add_paragraph(style="List Bullet")
    fmt(b, 0, 0, line=12.8, left=11, hanging=11)
    mark_size(b, 10)
    run(b, "Trained over 200 software developers in HTML, CSS, PHP, WordPress, and MySQL through practical lab sessions.", size=10, color=BODY)

    heading("Education", before=16, key="education")
    spacer(doc, 4, keep_next=True)
    border = {e: (CARD, 6) for e in ("top", "left", "bottom", "right")}
    t = make_table(doc, [CONTENT_TW])
    row_cant_split(t.rows[0])
    c = t.rows[0].cells[0]
    cell_props(c, fill=BG, borders=border, margins=(110 * G, 160, 110 * G, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 1.5, line=14)
    tab_right(p, CONTENT_TW - 320)
    run(p, "Bachelor of Science (B.Sc.) in Computer Science & Engineering", size=10.5, color=DARK, bold=True, font=HEADF)
    run(p, "\t2010 \u2013 2014", size=9.5, color=MUTED)
    p = c.add_paragraph()
    fmt(p, 0, 0, line=13)
    run(p, "Dhaka International University (DIU)", size=10, color=BLUE, bold=True)

    heading("Training & Professional Certifications", before=16, key="training")
    spacer(doc, 4, keep_next=True)
    gap = 160
    wd = (CONTENT_TW - gap) // 2
    t = make_table(doc, [wd, gap, wd])
    row_cant_split(t.rows[0])
    for ci, (title, sub) in ((0, ("Web Application Development with PHP", "BASIS (Bangladesh Association of Software & Information Services) \u2014 2015")),
                             (2, ("Web Design & Development", "Daffodil Institute of Information Technology (DIIT) \u2014 2012"))):
        cell = t.rows[0].cells[ci]
        cell_props(cell, borders=border, margins=(100 * G, 130, 100 * G, 130), width=wd)
        p = cell.paragraphs[0]
        fmt(p, 0, 1.5, line=12.5)
        run(p, title, size=10, color=DARK, bold=True)
        p = cell.add_paragraph()
        fmt(p, 0, 0, line=12)
        run(p, sub, size=9.5, color=MUTED)
    pp = t.rows[0].cells[1].paragraphs[0]
    fmt(pp, 0, 0, line=8, scale=False)
    mark_size(pp, 2)
    spacer(doc, 6 * g)

    t = make_table(doc, [CONTENT_TW])
    row_cant_split(t.rows[0])
    c = t.rows[0].cells[0]
    cell_props(c, fill=TAG, margins=(90 * G, 160, 90 * G, 160), width=CONTENT_TW)
    p = c.paragraphs[0]
    fmt(p, 0, 0, line=13)
    run(p, "Languages: ", size=10, color=DARK, bold=True)
    run(p, "Bangla (Native) | English (Proficiency)", size=10, color=INK)

    # final paragraph after the last table must exist; keep it microscopic
    last = doc.paragraphs[-1]
    if last.text.strip() or last._p.getnext() is not None and last._p.getnext().tag == qn("w:tbl"):
        pass
    endp = doc.add_paragraph()
    fmt(endp, 0, 0, line=1, scale=False)
    mark_size(endp, 1)

    footer_setup(doc.sections[0])
    if small:
        rescale_body(doc, 0.965)

    # widow control via non-breaking space between last two words
    for p in [Paragraph(e_, doc._body) for e_ in doc.element.body.iter(qn("w:p"))]:
        txt = p.text.strip()
        for (prevw, wd_) in nbsp:
            if txt.endswith(prevw + " " + wd_):
                join_last_two_words(p)
    fix_order(doc.element)
    fix_order(doc.sections[0].footer._element)
    return doc


# ----------------------------------------------------------------------------- measuring
class Measurer:
    def __init__(self):
        self.app = w.Dispatch("Word.Application")
        self.app.Visible = False
        self.n = 0

    def measure(self, doc, keep=None):
        self.n += 1
        docx = TMP / f"_cv{self.n}.docx"
        pdfp = TMP / f"_cv{self.n}.pdf"
        doc.save(docx)
        d = self.app.Documents.Open(str(docx), ReadOnly=True)
        pages = d.ComputeStatistics(2)
        d.SaveAs2(str(pdfp), 17)
        d.Close(False)
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
            first_heads = [ln[4] for ln in lines if ln[0] < 120][:2]
            widows = []
            for i in range(1, len(lines)):
                y0, y1, x0, x1, txt = lines[i]
                py0, py1, px0, px1, ptxt = lines[i - 1]
                if (y0 < 792 and len(txt.split()) == 1 and len(txt) > 3 and 56 < x0 < 80
                        and 0 < y0 - py0 < 16 and px1 > 440 and abs(px0 - x0) < 14):
                    widows.append((ptxt.split()[-1], txt))
            info.append(dict(bottom=bottom, widows=widows, first=first_heads))
        if keep:
            pdf.save(keep)
        pdf.close()
        for f in (docx,) + (() if keep else (pdfp,)):
            try:
                os.remove(f)
            except OSError:
                pass
        return pages, info

    def close(self):
        self.app.Quit()


def main(final=True):
    m = Measurer()
    nbsp = []
    try:
        pages = None
        for rnd in range(4):
            pages, info = m.measure(build(1.0, 1.0, nbsp, small=True))
            widows = [x for pg in info for x in pg["widows"]]
            print("round", rnd, "pages", pages, "bottoms", [round(i["bottom"]) for i in info], "widows", widows)
            new = [x for x in widows if x not in nbsp]
            if not new:
                break
            nbsp += new
        print("natural layout: pages", pages)
        for i, pg in enumerate(info):
            print("  page", i + 1, "starts with", pg["first"], "bottom", round(pg["bottom"]))
        target = 4

        def params(s):
            return 1 + (s - 1) * 1.0, 1 + (s - 1) * 0.35

        def ok(s):
            g, k = params(s)
            pg, inf = m.measure(build(g, k, nbsp, small=True))
            return pg == target and inf[-1]["bottom"] <= BOTTOM_LIMIT - 2, inf

        lo, hi = 0.5, 1.6
        for _ in range(8):
            mid = (lo + hi) / 2
            good, _inf = ok(mid)
            if good:
                lo = mid
            else:
                hi = mid
        print("scale", lo)
        # widow re-check at final scale
        for _ in range(3):
            g, k = params(lo)
            pg, inf = m.measure(build(g, k, nbsp, small=True))
            wd = [x for p_ in inf for x in p_["widows"] if x not in nbsp]
            print("final check pages", pg, "bottoms", [round(i["bottom"]) for i in inf], "new widows", wd)
            if not wd:
                break
            nbsp += wd
        g, k = params(lo)
        doc = build(g, k, nbsp, small=True)
        out = OUT
        try:
            doc.save(out)
        except PermissionError:
            out = OUT.with_name(OUT.stem + "_v2.docx")
            doc.save(out)
        d = m.app.Documents.Open(str(out), ReadOnly=True)
        pdf_out = out.with_suffix(".pdf")
        d.SaveAs2(str(pdf_out), 17)
        print("pages", d.ComputeStatistics(2))
        d.Close(False)
        print("saved", out, pdf_out)
    finally:
        m.close()
    return nbsp


if __name__ == "__main__":
    main()
