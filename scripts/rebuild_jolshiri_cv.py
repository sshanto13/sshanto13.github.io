"""Tailored CV for Jolshiri Abashon — Full Stack Web Developer (Web, Mobile & ERP)."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io")
OUT = ROOT / "public" / "files" / "Sumon_Ray_Full_Stack_Web_Developer_Jolshiri_2026.docx"
PHOTO = ROOT / "public" / "img" / "ss-linkedin.png"

ACCENT = RGBColor(0x0F, 0x76, 0x6E)
INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)

BODY = 9.5
SMALL = 9


def set_run(run, *, size=BODY, bold=False, color=INK, italic=False):
    run.font.name = "Calibri"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def spacing(p, before=0, after=2, line=1.08):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "12"), ("w:space", "3"), ("w:color", "0F766E")):
        bottom.set(qn(k), v)
    pBdr.append(bottom)
    pPr.append(pBdr)


def no_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tcPr.append(borders)


def keep_with_next(p):
    p.paragraph_format.keep_with_next = True


def heading(doc, text):
    p = doc.add_paragraph()
    spacing(p, before=9, after=3, line=1.0)
    set_run(p.add_run(text.upper()), size=11, bold=True, color=ACCENT)
    bottom_border(p)
    keep_with_next(p)


def body(doc, text, *, size=BODY, bold=False, before=0, after=2, color=INK, italic=False):
    p = doc.add_paragraph()
    spacing(p, before=before, after=after, line=1.1)
    set_run(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    spacing(p, before=0, after=1, line=1.08)
    set_run(p.add_run(text), size=BODY)


def skill_line(doc, label, items):
    p = doc.add_paragraph()
    spacing(p, before=1, after=1, line=1.08)
    set_run(p.add_run(f"{label}: "), size=BODY, bold=True)
    set_run(p.add_run(items), size=BODY)


def role(doc, company, meta, summary, bullets, tools):
    p = body(doc, company, size=10.5, bold=True, before=5, after=0)
    keep_with_next(p)
    p = body(doc, meta, size=SMALL, color=MUTED, after=1)
    keep_with_next(p)
    if summary:
        p = body(doc, summary, size=SMALL, italic=True, after=1)
        keep_with_next(p)
    for b in bullets:
        bullet(doc, b)
    if tools:
        p = doc.add_paragraph()
        spacing(p, before=1, after=1, line=1.05)
        set_run(p.add_run("Technologies: "), size=SMALL, bold=True, color=MUTED)
        set_run(p.add_run(tools), size=SMALL, color=MUTED)


def project(doc, title, stack, text):
    p = doc.add_paragraph()
    spacing(p, before=3, after=0, line=1.05)
    set_run(p.add_run(title), size=BODY, bold=True)
    keep_with_next(p)
    p = doc.add_paragraph()
    spacing(p, before=0, after=0, line=1.05)
    set_run(p.add_run(stack), size=SMALL, italic=True, color=MUTED)
    keep_with_next(p)
    body(doc, text, size=SMALL, after=1)


def header(doc):
    table = doc.add_table(rows=1, cols=2)
    text_cell, photo_cell = table.cell(0, 0), table.cell(0, 1)
    no_borders(text_cell)
    no_borders(photo_cell)
    text_cell.width = Inches(5.55)
    photo_cell.width = Inches(1.25)

    p = text_cell.paragraphs[0]
    spacing(p, after=1, line=1.0)
    set_run(p.add_run("SUMON RAY"), size=18, bold=True)

    p = text_cell.add_paragraph()
    spacing(p, after=3, line=1.05)
    set_run(
        p.add_run("Full Stack Web Developer  |  PHP, Laravel, ERP & E-commerce"),
        size=11,
        bold=True,
        color=ACCENT,
    )

    p = text_cell.add_paragraph()
    spacing(p, after=4, line=1.08)
    set_run(
        p.add_run(
            "8+ years building and maintaining websites, business applications, APIs, and "
            "admin systems — including e-commerce, ERP/POS, property operations, and government ICT."
        ),
        size=SMALL,
        color=MUTED,
    )

    p = text_cell.add_paragraph()
    spacing(p, after=1, line=1.05)
    set_run(p.add_run("Dhaka, Bangladesh  |  sumancsediu@gmail.com  |  +880 1952 499680"), size=SMALL)

    p = text_cell.add_paragraph()
    spacing(p, after=1, line=1.05)
    set_run(
        p.add_run(
            "LinkedIn: linkedin.com/in/sshanto13  |  Portfolio: sshanto13.github.io  |  "
            "GitHub: github.com/sshanto13"
        ),
        size=SMALL,
        color=MUTED,
    )

    p = photo_cell.paragraphs[0]
    spacing(p, after=0, line=1.0)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if PHOTO.exists():
        p.add_run().add_picture(str(PHOTO), width=Inches(1.1))


def build():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.45)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    header(doc)

    heading(doc, "Professional Profile")
    body(
        doc,
        "Full stack web developer with a B.Sc. in Computer Science & Engineering and 8+ years "
        "building, maintaining, and supporting production websites and business systems. Core "
        "stack is PHP and Laravel, with MySQL and PostgreSQL, JavaScript, and REST APIs, plus "
        "responsive front ends in Vue.js, React, and Angular.",
        after=2,
    )
    body(
        doc,
        "The closest match to this role is four years at PaiMart (JS Enterprise), where I worked "
        "across an e-commerce platform and ERP/POS: products, inventory, orders, customers, "
        "payments, billing, employee records, reporting, and admin workflows. I currently build "
        "property, booking, and operations tools for a travel company, and earlier delivered a "
        "government information system (NGTE Welfare Trust) covering applications, tracking, "
        "documents, and administrative workflows.",
        after=2,
    )
    body(
        doc,
        "I also handle the operational side of delivery: Git version control, domain and DNS, "
        "SSL, Linux VPS and Cloudflare, deployments, database care and backups, performance "
        "fixes, and production troubleshooting. I work directly with business users to turn "
        "manual processes into working software, and I keep technical and deployment notes.",
        after=1,
    )

    heading(doc, "Technical Skills")
    skill_line(
        doc,
        "Backend",
        "PHP, Laravel, REST API development and integration, Node.js, JavaScript, TypeScript",
    )
    skill_line(
        doc,
        "Frontend",
        "HTML5, CSS3, JavaScript, Vue.js, React, Angular, jQuery, AJAX, Bootstrap, Tailwind CSS, Inertia.js",
    )
    skill_line(
        doc,
        "Database",
        "MySQL, PostgreSQL, SQL, database design, query optimization, data modeling, Redis, backups",
    )
    skill_line(
        doc,
        "Business systems",
        "E-commerce, ERP, POS, inventory, orders, sales and billing, booking, property operations, dashboards and reports",
    )
    skill_line(
        doc,
        "Hosting & deployment",
        "Domain and DNS, SSL, Linux VPS, Cloudflare, AWS EC2/RDS/SES, DigitalOcean, Vultr, Docker, CI/CD",
    )
    skill_line(
        doc,
        "Security & delivery",
        "Application and web security practices, access control, Git/GitHub, code review, API testing, technical documentation",
    )
    skill_line(
        doc,
        "CMS",
        "WordPress, WooCommerce, Shopify, WordPress REST API",
    )

    heading(doc, "Professional Experience")

    role(
        doc,
        "DESTINIQUE TRAVEL LLC",
        "Software Engineer  |  Dhaka, Bangladesh  |  Feb 2025 – Present",
        "Web applications, APIs, and operations tools for a travel and property team.",
        [
            "Build and maintain systems for property management, booking, promotions, visitor tracking, and advertising.",
            "Develop REST APIs and admin dashboards used daily by the operations team.",
            "Contribute to the Angular customer-facing website and PHP/Laravel backend integrations.",
            "Tune slow database queries and fix production bugs, integration failures, and reliability issues.",
            "Manage deployment on Cloudflare and Linux VPS, including DNS, SSL, and routine maintenance, and keep database backups.",
            "Work with business stakeholders to turn requirements into working software and document the result.",
        ],
        "PHP, Laravel, MySQL, Angular, JavaScript, REST API, Git, Cloudflare, VPS",
    )

    role(
        doc,
        "PAIMART BY PAINEST — JS ENTERPRISE LTD.",
        "Assistant Manager, Engineering  |  Dhaka, Bangladesh  |  Jan 2021 – Jan 2025",
        "Four years on an e-commerce, distribution, and ERP/POS platform.",
        [
            "Built and maintained systems for products, inventory, orders, suppliers, customers, payments, invoicing, and billing.",
            "Supported wholesale/retail, B2B distribution, and field-sales workflows, plus ERP/POS and back-office operations.",
            "Covered employee management, operational reporting, and internal admin tools used by sales and operations staff.",
            "Exposed REST APIs for the storefront and for mobile clients, and worked with other teams on integrations.",
            "Handled releases, performance fixes, and production support as order volume and users grew.",
            "Helped replace manual operational steps with digital workflows inside the business platform.",
        ],
        "PHP, Laravel, MySQL, JavaScript, REST API, jQuery, Git, VPS",
    )

    role(
        doc,
        "ADN TELECOM PLC",
        "Assistant Manager, Engineering  |  Gulshan, Dhaka, Bangladesh  |  Oct 2019 – Jan 2021",
        "Enterprise web applications and a government ICT system.",
        [
            "Built Laravel backends and REST APIs for client and enterprise applications.",
            "Delivered the NGTE Welfare Trust system: application processing, tracking, document generation, and administrative workflows.",
            "Handled database schema work, testing, deployment support, and production issue resolution.",
        ],
        "PHP, Laravel, MySQL, JavaScript, REST API, Bootstrap, Git",
    )

    role(
        doc,
        "NYBSYS INC.",
        "Software Engineer  |  Mirpur, Dhaka, Bangladesh  |  Dec 2017 – Sep 2019",
        "",
        [
            "Built responsive web interfaces and integrated them with backend APIs.",
            "Worked with distributed teams on web deliverables and on features consumed by mobile clients.",
            "Turned designs and tickets into production screens and supported releases.",
        ],
        "PHP, JavaScript, HTML5, CSS3, API integration, Git",
    )

    role(
        doc,
        "ICT DIVISION — LEARNING & EARNING DEVELOPMENT PROJECT",
        "Trainer  |  Dhaka, Bangladesh  |  Dec 2016 – Apr 2017",
        "",
        [
            "Trained approximately 200 learners in HTML, CSS, PHP, WordPress, and MySQL through hands-on lab sessions.",
        ],
        "HTML, CSS, PHP, WordPress, MySQL",
    )

    heading(doc, "Selected Projects")
    project(
        doc,
        "PaiMart — Wholesale E-commerce  ·  paimart.painest.com",
        "Laravel  |  PHP  |  MySQL  |  JavaScript  |  REST API",
        "Storefront and back office for products, inventory, orders, payments, customers, ERP/POS, and reporting.",
    )
    project(
        doc,
        "JS ERP — Sales, Billing & Inventory  ·  js-erp.paitechno.com",
        "Laravel  |  PHP  |  MySQL  |  JavaScript",
        "Business system for sales, invoicing, billing, inventory, employee records, and day-to-day operations.",
    )
    project(
        doc,
        "E-commerce Operations — Back Office  ·  operation.paikaree.com.bd",
        "Laravel  |  PHP  |  MySQL  |  JavaScript",
        "Internal admin system for product, order, and inventory operations.",
    )
    project(
        doc,
        "Destinique Travel Platform  ·  destinique.com",
        "Angular  |  PHP  |  Laravel  |  REST API  |  MySQL  |  Cloudflare",
        "Customer website for booking, promotions, and property experiences, with Laravel APIs behind the Angular front end.",
    )
    project(
        doc,
        "Destinique Business Dashboard  ·  quote.destinique.com/destin/dashboard",
        "PHP  |  MySQL  |  JavaScript  |  REST API  |  Cloudflare",
        "Internal dashboard for visitor tracking, promotions, property management, booking, and operational reporting.",
    )
    project(
        doc,
        "NGTE Welfare Trust  ·  apps.ngte-welfaretrust.gov.bd",
        "Laravel  |  PHP  |  MySQL  |  JavaScript",
        "Government system for application processing, tracking, documents, and administrative workflows.",
    )

    heading(doc, "Education")
    p = body(doc, "B.Sc. in Computer Science & Engineering", size=10.5, bold=True, before=1, after=0)
    keep_with_next(p)
    body(doc, "Dhaka International University  |  2010 – 2015", size=SMALL, color=MUTED)

    heading(doc, "Professional Training")
    p = body(doc, "Web Application Development with PHP", size=10, bold=True, before=1, after=0)
    keep_with_next(p)
    body(
        doc,
        "Bangladesh Association of Software and Information Services (BASIS)  |  2015",
        size=SMALL,
        color=MUTED,
    )
    p = body(doc, "Web Design & Development", size=10, bold=True, before=2, after=0)
    keep_with_next(p)
    body(doc, "Daffodil Institute of IT (DIIT)  |  2012", size=SMALL, color=MUTED)

    heading(doc, "Languages")
    body(doc, "Bangla (native)  |  English (professional working proficiency)")

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_v2.docx")
        doc.save(alt)
        print("locked; saved", alt)


if __name__ == "__main__":
    build()
