"""Tailored CV for The Asia Foundation Bangladesh — IT Officer."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io")
OUT = ROOT / "public" / "files" / "Sumon_Ray_IT_Officer_Asia_Foundation_2026_final.docx"
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


def spacing(p, before=0, after=2, line=1.1):
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
    spacing(p, before=10, after=4, line=1.0)
    set_run(p.add_run(text.upper()), size=11, bold=True, color=ACCENT)
    bottom_border(p)
    keep_with_next(p)


def body(doc, text, *, size=BODY, bold=False, before=0, after=2, color=INK, italic=False):
    p = doc.add_paragraph()
    spacing(p, before=before, after=after, line=1.12)
    set_run(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    spacing(p, before=0, after=1.5, line=1.1)
    set_run(p.add_run(text), size=BODY)


def labeled(doc, label, text, *, as_bullet=False):
    p = doc.add_paragraph(style="List Bullet") if as_bullet else doc.add_paragraph()
    spacing(p, before=0 if as_bullet else 1, after=2, line=1.1)
    set_run(p.add_run(f"{label}: "), size=BODY, bold=True)
    set_run(p.add_run(text), size=BODY)


def role(doc, company, meta, summary, bullets, tools):
    p = body(doc, company, size=10.5, bold=True, before=6, after=0)
    keep_with_next(p)
    p = body(doc, meta, size=SMALL, color=MUTED, after=2)
    keep_with_next(p)
    if summary:
        p = body(doc, summary, size=SMALL, italic=True, after=2)
        keep_with_next(p)
    for b in bullets:
        bullet(doc, b)
    if tools:
        p = doc.add_paragraph()
        spacing(p, before=1, after=2, line=1.1)
        set_run(p.add_run("Tools: "), size=SMALL, bold=True, color=MUTED)
        set_run(p.add_run(tools), size=SMALL, color=MUTED)


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
        p.add_run("IT Systems & Support Professional | Infrastructure, Security & Applications"),
        size=11,
        bold=True,
        color=ACCENT,
    )

    p = text_cell.add_paragraph()
    spacing(p, after=4, line=1.1)
    set_run(
        p.add_run(
            "8+ years keeping business IT systems stable, secure, and usable: servers, cloud/VPS, "
            "DNS/SSL, access control, incident troubleshooting, user support, and staff training."
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
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    header(doc)

    heading(doc, "Professional Profile")
    body(
        doc,
        "IT professional with a B.Sc. in Computer Science & Engineering and 8+ years of hands-on "
        "experience across software engineering, business systems, technical support, server/cloud "
        "environments, and ICT operations. My work has involved keeping business-critical systems "
        "running for staff and users, including diagnosing and resolving incidents, supporting "
        "production environments, managing Linux/VPS infrastructure, handling DNS, SSL and "
        "Cloudflare security settings, managing application access, and coordinating with "
        "stakeholders and service providers when issues require escalation.",
        after=3,
    )
    body(
        doc,
        "I am security-conscious and process-driven, with experience following access controls, "
        "deployment procedures, documentation practices, and backup/recovery processes. I have "
        "trained approximately 200 learners through a Government of Bangladesh ICT programme and "
        "regularly support non-technical users by explaining technical issues clearly and "
        "practically. I communicate effectively in English and Bangla and am comfortable working "
        "independently, with distributed teams, and providing occasional after-hours support when "
        "urgent issues arise.",
        after=2,
    )

    heading(doc, "Core Competencies")
    for label, text in [
        (
            "IT Support & Troubleshooting",
            "Level 1–2 incident diagnosis, root-cause analysis, log review, issue tracking and "
            "escalation, production support, user guidance, troubleshooting of business "
            "applications and technical issues",
        ),
        (
            "Desktop & User Support",
            "Windows 10/11 installation and troubleshooting, Microsoft 365 user-level support, "
            "basic LAN/Wi-Fi troubleshooting, printer and peripheral troubleshooting, remote "
            "technical support, email account setup and administration",
        ),
        (
            "Endpoint & Asset Management",
            "Antivirus management, disk encryption and device management tools (BitLocker, "
            "Intune), IT asset and device inventory, user account onboarding and offboarding",
        ),
        (
            "Networking & Infrastructure",
            "Basic IP concepts, DNS management, SSL/TLS certificates, Cloudflare DNS and security "
            "settings, Linux servers, cloud/VPS administration, hosting environments, deployment "
            "and maintenance, ISP/provider coordination",
        ),
        (
            "Security & Access Control",
            "Role-based access control, user permissions, least-privilege access, secure "
            "configuration, server/application updates and patching, incident reporting, backup "
            "and recovery, password and MFA support, staff security awareness",
        ),
        (
            "Service Management & Reporting",
            "Issue and service-request tracking, Jira/Trello, technical documentation, user guides, "
            "operational dashboards, periodic reporting, issue escalation, vendor and "
            "service-provider liaison, process improvement",
        ),
        (
            "Training & Communication",
            "User guidance, staff onboarding support, IT orientation, hands-on technical training "
            "for approximately 200 learners, clear written and verbal communication in English "
            "and Bangla",
        ),
        (
            "Applications & Systems",
            "ERP, POS, inventory, sales and billing systems, booking and property-related systems, "
            "web applications, MySQL/PostgreSQL, PHP/Laravel, JavaScript, REST APIs, Git",
        ),
    ]:
        labeled(doc, label, text)

    heading(doc, "Professional Experience")

    role(
        doc,
        "DESTINIQUE TRAVEL LLC — US-based Travel Company",
        "Software Engineer  |  Dhaka, Bangladesh  |  Feb 2025 – Present",
        "Supporting the systems, infrastructure, and internal tools a remote business team relies on daily.",
        [
            "Provide day-to-day technical support to business users on internal systems used for "
            "bookings, property records, promotions, and reporting; resolve issues and guide users.",
            "Diagnose and resolve production incidents involving application errors, slow queries, "
            "integration failures, and system issues, tracking problems through to resolution and "
            "follow-up.",
            "Manage Cloudflare and VPS infrastructure, including DNS, SSL, security settings, "
            "deployments, and routine maintenance to keep services available and secure.",
            "Set up and verify server and database backups, and restore services when failures "
            "occur.",
            "Liaise with hosting, cloud, and service providers, including Cloudflare and VPS "
            "vendors, to resolve incidents and service issues.",
            "Support analytics and tracking integrations and maintain operational dashboards used "
            "for monitoring and reporting.",
            "Coordinate with stakeholders across locations to gather requirements, prioritize "
            "requests, resolve issues, and communicate technical status clearly.",
            "Maintain technical systems and documentation while following controlled deployment "
            "and access practices.",
        ],
        "Cloudflare, Linux VPS, DNS/SSL, MySQL, PHP/Laravel, Angular, REST APIs, Git, dashboards",
    )

    role(
        doc,
        "PAIMART BY PAINEST — JS ENTERPRISE LTD.",
        "Assistant Manager, Engineering  |  Dhaka, Bangladesh  |  Jan 2021 – Jan 2025",
        "Responsible for technology supporting a multi-team e-commerce, distribution, and ERP/POS operation.",
        [
            "Kept business-critical systems for sales, inventory, orders, billing, payments, and "
            "back-office operations stable as data volume and user numbers grew.",
            "Handled incident response and production troubleshooting for operations, sales, and "
            "back-office staff; escalated and coordinated fixes across internal teams.",
            "Managed application releases, server/application updates, performance improvements, "
            "and hosting on VPS infrastructure, following controlled deployment practices.",
            "Implemented user roles and access permissions in business applications so staff could "
            "access only the data and functions required for their responsibilities.",
            "Handled staff onboarding and offboarding, including creating and removing user and "
            "email accounts and granting or revoking system access.",
            "Maintained an inventory of IT assets and devices, and supported endpoint security "
            "with antivirus and device-management tools.",
            "Supported staff with password resets and MFA, and briefed users on phishing and "
            "safe computing practices.",
            "Produced reports and dashboards for management monitoring and supported users across "
            "departments with day-to-day system questions.",
            "Investigated recurring technical problems and implemented improvements to reduce "
            "operational issues and improve system reliability.",
        ],
        "Linux VPS, MySQL, PHP/Laravel, JavaScript, REST APIs, ERP/POS, Git",
    )

    role(
        doc,
        "ADN TELECOM PLC",
        "Assistant Manager, Engineering  |  Gulshan, Dhaka, Bangladesh  |  Oct 2019 – Jan 2021",
        "Worked on enterprise applications and government ICT systems at a telecom and ICT service provider.",
        [
            "Delivered the NGTE Welfare Trust government information system covering application "
            "processing, tracking, document generation, and administrative workflows.",
            "Supported deployment, testing, troubleshooting, and production issue resolution for "
            "enterprise and government ICT applications.",
            "Worked with structured administrative records where access control, data accuracy, and "
            "reliable handling of information were essential.",
            "Coordinated with project stakeholders to investigate issues, implement fixes, and "
            "maintain reliable application operations.",
        ],
        "PHP/Laravel, MySQL, JavaScript, REST APIs, Bootstrap",
    )

    role(
        doc,
        "NYBSYS INC.",
        "Software Engineer  |  Mirpur, Dhaka, Bangladesh  |  Dec 2017 – Sep 2019",
        "",
        [
            "Built and integrated web applications while working with distributed and international "
            "teams across web, mobile, IoT, LTE, and telecommunications-related projects.",
            "Supported testing, debugging, integration, troubleshooting, and delivery of production "
            "releases.",
            "Investigated technical issues and coordinated with team members to resolve application "
            "and integration problems.",
        ],
        "PHP, JavaScript, HTML, CSS, REST/API integration",
    )

    role(
        doc,
        "ICT DIVISION — LEARNING & EARNING DEVELOPMENT PROJECT",
        "Trainer  |  Dhaka, Bangladesh  |  Dec 2016 – Apr 2017",
        "",
        [
            "Trained approximately 200 learners through hands-on laboratory sessions on web "
            "technologies and databases under a Government of Bangladesh ICT programme.",
            "Explained technical concepts clearly to non-specialist audiences and provided "
            "practical guidance during hands-on sessions.",
            "Developed communication and training skills relevant to user induction, technical "
            "orientation, and staff awareness activities.",
        ],
        "",
    )

    heading(doc, "Education")
    p = body(doc, "B.Sc. in Computer Science & Engineering", size=10.5, bold=True, before=1, after=0)
    keep_with_next(p)
    body(doc, "Dhaka International University  |  2010 – 2015", size=SMALL, color=MUTED)

    heading(doc, "Professional Training")
    p = body(doc, "Web Application Development with PHP", size=10, bold=True, before=1, after=0)
    keep_with_next(p)
    body(doc, "Bangladesh Association of Software and Information Services (BASIS)  |  2015", size=SMALL, color=MUTED)
    p = body(doc, "Web Design & Development", size=10, bold=True, before=3, after=0)
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
