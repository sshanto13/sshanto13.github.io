"""Rebuild Portfolio CV to align with current website content."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files\Sumon_Ray_Software_Engineer_Portfolio_2026.docx")
OUT_UPDATED = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files\Sumon_Ray_Software_Engineer_Portfolio_2026_updated.docx")
PHOTO = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\img\ss-linkedin.png")

ACCENT = RGBColor(0x0F, 0x76, 0x6E)
INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)


def set_run(run, *, size=10, bold=False, color=INK, name="Calibri"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def set_para_spacing(p, before=0, after=4, line=1.15):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def add_bottom_border(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "0F766E")
    pBdr.append(bottom)
    pPr.append(pBdr)


def heading(doc, text):
    p = doc.add_paragraph()
    set_para_spacing(p, before=10, after=4, line=1.1)
    r = p.add_run(text.upper())
    set_run(r, size=11, bold=True, color=ACCENT)
    add_bottom_border(p)
    return p


def body(doc, text, *, size=10, bold=False, before=0, after=3):
    p = doc.add_paragraph()
    set_para_spacing(p, before=before, after=after, line=1.15)
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=INK)
    return p


def muted(doc, text, *, size=9.5, before=0, after=2):
    p = doc.add_paragraph()
    set_para_spacing(p, before=before, after=after, line=1.1)
    r = p.add_run(text)
    set_run(r, size=size, bold=False, color=MUTED)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    set_para_spacing(p, before=0, after=2, line=1.12)
    # clear default run and set ours
    if p.runs:
        p.runs[0].text = text
        set_run(p.runs[0], size=9.5, color=INK)
    else:
        r = p.add_run(text)
        set_run(r, size=9.5, color=INK)
    return p


def skill_line(doc, label, items):
    p = doc.add_paragraph()
    set_para_spacing(p, before=1, after=2, line=1.12)
    r1 = p.add_run(f"{label}: ")
    set_run(r1, size=9.5, bold=True, color=INK)
    r2 = p.add_run(items)
    set_run(r2, size=9.5, bold=False, color=INK)
    return p


def set_cell_border(cell, **kwargs):
    """Remove/set cell borders. kwargs: top/left/bottom/right -> None to remove."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        tcBorders.append(element)
    tcPr.append(tcBorders)


def build():
    doc = Document()

    # Margins — compact but readable
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    # Header: identity left + photo right
    header = doc.add_table(rows=1, cols=2)
    header.autofit = True
    text_cell = header.cell(0, 0)
    photo_cell = header.cell(0, 1)
    set_cell_border(photo_cell)
    set_cell_border(text_cell)

    text_cell.width = Inches(5.6)
    photo_cell.width = Inches(1.25)

    # Name / title / contact in left cell
    p_name = text_cell.paragraphs[0]
    set_para_spacing(p_name, before=2, after=2, line=1.0)
    r = p_name.add_run("SUMON RAY")
    set_run(r, size=18, bold=True, color=INK)

    p_title = text_cell.add_paragraph()
    set_para_spacing(p_title, before=0, after=2, line=1.05)
    r = p_title.add_run("Software Engineer | Solutions Developer")
    set_run(r, size=11, bold=True, color=ACCENT)

    p_tag = text_cell.add_paragraph()
    set_para_spacing(p_tag, before=0, after=4, line=1.1)
    r = p_tag.add_run(
        "I build software people at work can actually run — web apps, APIs, "
        "business systems, and automation."
    )
    set_run(r, size=9.5, color=MUTED)

    p_contact = text_cell.add_paragraph()
    set_para_spacing(p_contact, before=0, after=1, line=1.1)
    r = p_contact.add_run(
        "Dhaka, Bangladesh  ·  sumancsediu@gmail.com  ·  +880 1952 499680"
    )
    set_run(r, size=9, color=INK)

    p_links = text_cell.add_paragraph()
    set_para_spacing(p_links, before=0, after=2, line=1.1)
    r = p_links.add_run(
        "LinkedIn: linkedin.com/in/sshanto13  ·  GitHub: github.com/sshanto13  ·  "
        "Portfolio: sshanto13.github.io"
    )
    set_run(r, size=9, color=MUTED)

    # Photo on the right
    p_photo = photo_cell.paragraphs[0]
    set_para_spacing(p_photo, before=0, after=0, line=1.0)
    p_photo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if PHOTO.exists():
        run = p_photo.add_run()
        run.add_picture(str(PHOTO), width=Inches(1.15))

    # spacer under header
    spacer = doc.add_paragraph()
    set_para_spacing(spacer, before=2, after=2, line=1.0)

    # Profile — aligned with About section (professional CV tone)
    heading(doc, "Professional Profile")
    body(
        doc,
        "Software Engineer with 8+ years of experience building practical software "
        "solutions for real-world business operations across software companies, "
        "enterprise systems, e-commerce, government ICT, and travel & property management.",
        after=3,
    )
    body(
        doc,
        "Strongest ground: PHP, Laravel, SQL, MySQL, PostgreSQL, JavaScript, and REST APIs, "
        "with experience across Angular, Vue, React, Node.js, Python, cloud/VPS infrastructure, "
        "Docker, reporting/dashboards, and AI-assisted development.",
        after=3,
    )
    body(
        doc,
        "4+ years with PaiMart (previously Paikaree.com.bd, backed by JS Enterprise) building "
        "across e-commerce, sales, distribution, inventory, orders, payments, ERP/POS, reporting, "
        "APIs, and business workflows. Currently at Destinique Travel LLC on business automation "
        "and digital solutions covering booking, property management, promotions, visitor tracking, "
        "APIs, dashboards, and operational workflows.",
        after=3,
    )
    body(
        doc,
        "Open to software engineering, full-stack, backend (PHP/Laravel and related), "
        "business applications, automation, ERP/ICT, and data/reporting roles.",
        after=4,
    )

    # Skills — match portfolio skill categories
    heading(doc, "Technical Skills")
    skill_line(
        doc,
        "Programming & Backend",
        "PHP, Laravel, Node.js, Python, Go, JavaScript, TypeScript, REST APIs, API Integration, Web Services",
    )
    skill_line(
        doc,
        "Business & Systems",
        "E-commerce, ERP, POS, Inventory, Sales & Billing, Booking Systems, Property Management, Process Automation",
    )
    skill_line(
        doc,
        "Database & Data",
        "MySQL, PostgreSQL, MongoDB, SQL, Database Design, Query Optimization, Data Modeling, Redis",
    )
    skill_line(
        doc,
        "Frontend",
        "Angular, Vue.js, React, JavaScript, TypeScript, jQuery, AJAX, HTML5, CSS3, Bootstrap, Tailwind CSS, Inertia.js",
    )
    skill_line(
        doc,
        "Cloud, DevOps & Infra",
        "AWS EC2 / RDS / SES, Cloudflare, DigitalOcean, Vultr, Cloud VPS, Docker, CI/CD, DNS, SSL, Linux",
    )
    skill_line(
        doc,
        "BI, Reporting & Analytics",
        "Power BI, Dashboards, Custom Reports, Google Analytics 4, Data Tracking, Conversion Tracking",
    )
    skill_line(
        doc,
        "AI-Assisted Development",
        "Cursor, Claude, AI-Assisted Coding, Rapid MVP Development, Debugging, Review",
    )
    skill_line(
        doc,
        "Engineering Practices",
        "Git, System Design, Clean Architecture, Code Review, API Testing, Debugging, Optimisation, Security Practices",
    )
    skill_line(
        doc,
        "CMS & SEO",
        "WordPress, WooCommerce, Shopify, Wix, WordPress REST API, SEO Integration",
    )
    skill_line(
        doc,
        "Delivery & Collaboration",
        "Requirements, Stakeholders, Teamwork, Planning, Delivery, Mentoring, Production Support, Coordination",
    )
    skill_line(
        doc,
        "Tools & Workflow",
        "VS Code, IntelliJ IDEA, Sublime Text, Jira, Slack, Trello, Agile / Scrum, Technical Documentation",
    )

    # Experience — portfolio facts, CV-ready wording
    heading(doc, "Professional Experience")

    # Destinique
    body(doc, "DESTINIQUE TRAVEL LLC", bold=True, before=4, after=1)
    muted(doc, "Software Engineer  ·  Dhaka, Bangladesh  ·  Feb 2025 – Present")
    body(
        doc,
        "Business automation, digital solutions, and APIs for a luxury travel / property team.",
        size=9.5,
        after=2,
    )
    for b in [
        "Build and maintain solutions for visitor tracking, advertising, promotions, property management, and booking.",
        "Develop REST APIs and operational dashboards used day to day by the business team.",
        "Contribute to the Angular customer-facing platform and PHP/Laravel backend integrations.",
        "Tune slow queries and investigate production bugs, integrations, and reliability issues.",
        "Work with business stakeholders to translate requirements into shippable software.",
        "Support deployment, Cloudflare/VPS infrastructure, analytics integration, and production maintenance.",
    ]:
        bullet(doc, b)
    muted(doc, "Technologies: PHP, Laravel, MySQL, Angular, JavaScript, REST API, Cloudflare, VPS", after=4)

    # PaiMart
    body(doc, "PAIMART BY PAINEST — JS ENTERPRISE LTD.", bold=True, before=2, after=1)
    muted(doc, "Assistant Manager, Engineering  ·  Dhaka, Bangladesh  ·  Jan 2021 – Jan 2025")
    body(
        doc,
        "4+ years building business automation and digital solutions across e-commerce, distribution, "
        "wholesale/retail operations, ERP/POS, and day-to-day business workflows.",
        size=9.5,
        after=2,
    )
    for b in [
        "Built and maintained systems across e-commerce, B2B distribution, wholesale/retail and field sales, products, inventory, orders, and suppliers.",
        "Covered customers, payments, invoicing, billing, employee management, back-office operations, ERP/POS, reporting, APIs, and business workflows.",
        "Developed a strong understanding of how a complete business platform supports sales, distribution, and operations—not only the software layer.",
        "Exposed APIs for storefront and mobile; collaborated with teams on integrations and production issues.",
        "Handled releases, performance fixes, and production support as volume grew.",
    ]:
        bullet(doc, b)
    muted(doc, "Technologies: PHP, Laravel, MySQL, JavaScript, REST API, jQuery, Git, VPS", after=4)

    # ADN
    body(doc, "ADN TELECOM PLC", bold=True, before=2, after=1)
    muted(doc, "Assistant Manager, Engineering  ·  Gulshan, Dhaka, Bangladesh  ·  Oct 2019 – Jan 2021")
    body(doc, "Enterprise web applications and a government ICT project while at ADN.", size=9.5, after=2)
    for b in [
        "Built Laravel backends and APIs for client and enterprise applications.",
        "Delivered work on NGTE Welfare Trust — applications, tracking, documents, and administrative flows.",
        "Handled schema work, testing, deployment support, and production issue resolution.",
    ]:
        bullet(doc, b)
    muted(doc, "Technologies: PHP, Laravel, MySQL, JavaScript, REST API, Bootstrap", after=4)

    # NYBSYS
    body(doc, "NYBSYS INC.", bold=True, before=2, after=1)
    muted(doc, "Software Engineer  ·  Mirpur, Dhaka, Bangladesh  ·  Dec 2017 – Sep 2019")
    body(doc, "Web UI and API work on projects with distributed teams.", size=9.5, after=2)
    for b in [
        "Built frontends and integrated them with backend APIs.",
        "Collaborated with remote teammates on web and mobile-related deliverables.",
        "Turned designs and tickets into working production screens.",
    ]:
        bullet(doc, b)
    muted(doc, "Technologies: PHP, JavaScript, HTML, CSS, API Integration", after=4)

    # ICT
    body(doc, "ICT DIVISION — LEARNING & EARNING DEVELOPMENT PROJECT", bold=True, before=2, after=1)
    muted(doc, "Trainer  ·  Dhaka, Bangladesh  ·  Dec 2016 – Apr 2017")
    for b in [
        "Trained approximately 200 learners in HTML, CSS, PHP, WordPress, and MySQL.",
        "Delivered hands-on lab sessions focused on practical web development skills.",
    ]:
        bullet(doc, b)
    muted(doc, "Technologies: HTML, CSS, PHP, WordPress, MySQL", after=4)

    # Projects — portfolio featured set
    heading(doc, "Selected Projects")

    projects = [
        (
            "Destinique Travel Platform",
            "https://destinique.com/",
            "Angular | PHP | REST API | MySQL | JavaScript | Bootstrap | Cloudflare",
            "Customer-facing travel and property platform covering booking flows, promotions, and property experiences. Contributed across Angular frontend and PHP/Laravel backend APIs used by visitors and the operations team.",
        ),
        (
            "Destinique Business Dashboard",
            "https://quote.destinique.com/destin/dashboard/",
            "PHP | MySQL | JavaScript | REST API | Cloudflare",
            "Internal business platform supporting visitor tracking, advertising, promotions, property management, booking, reporting, and operational workflows.",
        ),
        (
            "PaiMart — Wholesale E-commerce",
            "https://paimart.painest.com/",
            "Laravel | PHP | MySQL | JavaScript | REST API",
            "Wholesale e-commerce and business platform spanning storefront operations and back-office workflows across products, inventory, orders, payments, ERP/POS, reporting, and APIs.",
        ),
        (
            "JS ERP — Sales, Billing & Inventory",
            "https://js-erp.paitechno.com/",
            "Laravel | PHP | MySQL | JavaScript",
            "Business management solution supporting sales, invoicing, billing, inventory, employee, and operational processes.",
        ),
        (
            "E-commerce Operations — Back Office",
            "https://operation.paikaree.com.bd/",
            "Laravel | PHP | MySQL | JavaScript",
            "Internal back-office system supporting product, order, inventory, and e-commerce operations.",
        ),
        (
            "NGTE Welfare Trust",
            "https://apps.ngte-welfaretrust.gov.bd/",
            "Laravel | PHP | MySQL | JavaScript",
            "Government information and application management system supporting application processing, tracking, document generation, and administrative workflows.",
        ),
    ]

    for title_t, url, stack, desc in projects:
        p = doc.add_paragraph()
        set_para_spacing(p, before=3, after=1, line=1.1)
        r = p.add_run(title_t)
        set_run(r, size=10, bold=True, color=INK)
        r2 = p.add_run(f"  ({url})")
        set_run(r2, size=8.5, color=MUTED)
        muted(doc, stack, size=8.5, after=1)
        body(doc, desc, size=9.5, after=2)

    # Where most useful — from portfolio usefulFor
    heading(doc, "Where I Am Most Useful")
    for item in [
        "Backend and full-stack development where APIs, databases, and business logic matter.",
        "Business applications: e-commerce, ERP/POS, inventory, booking, and automation.",
        "Data, reporting, dashboards, and operational automation tied to real business needs.",
        "Cross-stack problem solving across backend, frontend, databases, APIs, and infrastructure.",
        "Production-focused engineering: shipping features, troubleshooting, optimization, and keeping systems reliable.",
    ]:
        bullet(doc, item)

    # Education
    heading(doc, "Education")
    body(doc, "B.Sc. in Computer Science & Engineering", bold=True, before=2, after=1)
    muted(doc, "Dhaka International University  ·  2010 – 2015")

    heading(doc, "Professional Training")
    body(doc, "Web Application Development with PHP", bold=True, before=2, after=1)
    muted(doc, "BITM, BASIS  ·  2015")
    body(doc, "Web Design & Development", bold=True, before=3, after=1)
    muted(doc, "Daffodil Institute of IT (DIIT)  ·  2012")

    import shutil

    targets = [
        OUT,
        OUT_UPDATED,
        OUT.with_name("Sumon_Ray_Software_Engineer_Portfolio_2026_with_photo.docx"),
        OUT.with_name("Sumon_Ray_Software_Engineer_Portfolio_2026_final.docx"),
    ]
    saved = None
    for path in targets:
        try:
            doc.save(path)
            saved = path
            print("saved", path.name)
            break
        except PermissionError:
            print("locked:", path.name)
            continue

    if saved is None:
        raise SystemExit("Could not save CV — close Word and retry")

    for path in targets:
        if path == saved:
            continue
        try:
            shutil.copy2(saved, path)
            print("copied to", path.name)
        except Exception as e:
            print("skip", path.name, ":", e)

    print("ACTIVE_FILE", saved.name)
    print("photo:", PHOTO.exists(), PHOTO)


if __name__ == "__main__":
    build()
