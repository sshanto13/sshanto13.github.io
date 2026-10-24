"""CV for ACI Godrej Agrovet — Lead, IT & Digitalization (plain-language version).

Built from the user's own corrected wording (titles, dates, Florida, Sysba,
engagement roles). Each employer/client gets a short factual "about" line.
Does not claim SAP, SD-WAN, IT audits, budgets, headcount or certifications.
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io")
OUT = ROOT / "public" / "files" / "Sumon_Ray_Lead_IT_Digitalization_ACI_Godrej_2026_final.docx"
PHOTO = ROOT / "public" / "img" / "ss-linkedin.png"

ACCENT = RGBColor(0x0F, 0x76, 0x6E)
INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)

BODY = 10
SMALL = 9.5


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
    spacing(p, before=10, after=3, line=1.0)
    set_run(p.add_run(text.upper()), size=11.5, bold=True, color=ACCENT)
    bottom_border(p)
    keep_with_next(p)


def para(doc, text, *, size=BODY, bold=False, before=0, after=4, color=INK, italic=False):
    p = doc.add_paragraph()
    spacing(p, before=before, after=after, line=1.12)
    set_run(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    spacing(p, before=0, after=1.5, line=1.1)
    set_run(p.add_run(text), size=BODY)


def skill_line(doc, label, text):
    p = doc.add_paragraph()
    spacing(p, before=1, after=2, line=1.1)
    set_run(p.add_run(f"{label}: "), size=BODY, bold=True)
    set_run(p.add_run(text), size=BODY)


def project_bullet(doc, name, detail):
    p = doc.add_paragraph(style="List Bullet")
    spacing(p, before=0, after=1, line=1.05)
    set_run(p.add_run(name + ": "), size=SMALL, bold=True)
    set_run(p.add_run(detail), size=SMALL)


def role(doc, company, meta, about, bullets, tools="", projects=None, extras=None,
         tools_label="Technologies & Infrastructure"):
    p = para(doc, company, size=11.5, bold=True, before=7, after=0)
    keep_with_next(p)
    p = para(doc, meta, size=SMALL, color=MUTED, after=1)
    keep_with_next(p)
    if about:
        p = para(doc, about, size=SMALL, italic=True, color=MUTED, after=2)
        keep_with_next(p)
    for b in bullets:
        bullet(doc, b)
    if projects:
        p = doc.add_paragraph()
        spacing(p, before=2, after=1, line=1.05)
        set_run(p.add_run("Key Projects & Platforms:"), size=SMALL, bold=True, color=ACCENT)
        keep_with_next(p)
        for name, detail in projects:
            project_bullet(doc, name, detail)
    for label, text in (extras or []):
        p = doc.add_paragraph()
        spacing(p, before=1, after=1, line=1.05)
        set_run(p.add_run(label + ": "), size=SMALL, bold=True, color=ACCENT)
        set_run(p.add_run(text), size=SMALL)
    if tools:
        p = doc.add_paragraph()
        spacing(p, before=1, after=1, line=1.05)
        set_run(p.add_run(tools_label + ": "), size=SMALL, bold=True, color=ACCENT)
        set_run(p.add_run(tools), size=SMALL)


def entry(doc, name, meta, about, detail):
    """Compact engagement: bold name, muted role/dates, about line, one detail line."""
    p = doc.add_paragraph()
    spacing(p, before=5, after=0, line=1.1)
    set_run(p.add_run(name), size=BODY + 0.5, bold=True)
    set_run(p.add_run("  |  " + meta), size=SMALL, color=MUTED)
    keep_with_next(p)
    if about:
        p = para(doc, about, size=SMALL, italic=True, color=MUTED, after=1)
        keep_with_next(p)
    if detail:
        para(doc, detail, after=2)


def project(doc, title, detail):
    p = doc.add_paragraph()
    spacing(p, before=1, after=1.5, line=1.1)
    set_run(p.add_run(title), size=BODY, bold=True)
    set_run(p.add_run("  —  " + detail), size=BODY)


def group_label(doc, text):
    p = para(doc, text, size=10.5, bold=True, before=5, after=1, color=ACCENT)
    keep_with_next(p)


def header(doc):
    table = doc.add_table(rows=1, cols=2)
    text_cell, photo_cell = table.cell(0, 0), table.cell(0, 1)
    no_borders(text_cell)
    no_borders(photo_cell)
    text_cell.width = Inches(5.55)
    photo_cell.width = Inches(1.25)

    p = text_cell.paragraphs[0]
    spacing(p, after=1, line=1.0)
    set_run(p.add_run("SUMON RAY"), size=22, bold=True)

    p = text_cell.add_paragraph()
    spacing(p, after=3, line=1.05)
    set_run(
        p.add_run("ERP, IT & Digitalization  |  Infrastructure, Security & Business Automation"),
        size=11.5,
        bold=True,
        color=ACCENT,
    )

    p = text_cell.add_paragraph()
    spacing(p, after=4, line=1.08)
    set_run(
        p.add_run(
            "IT infrastructure, ERP, and digital solutions since 2014, "
            "across Bangladesh, South Africa, the USA, Belgium, and Spain."
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
        p.add_run("linkedin.com/in/sshanto13  |  sshanto13.github.io  |  github.com/sshanto13"),
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
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    header(doc)

    # ----------------------------------------------------------- Introduction
    heading(doc, "Professional Introduction")
    para(
        doc,
        "IT and Digitalization professional with 10+ years of experience since 2014 across enterprise "
        "applications, ERP development, IT infrastructure, business automation, and digital solutions. "
        "Combines hands-on technical expertise with business-focused delivery, including requirements "
        "analysis, process mapping, system integration, implementation, and user adoption.",
        after=3,
    )
    skill_line(
        doc,
        "Leadership & Enterprise Applications",
        "Independently designed and developed JS ERP at JS Group, progressing from Deputy Manager to "
        "Technical Manager. At Destinique Travel LLC (USA), lead technical delivery and business automation "
        "across travel and property-management workflows. Consultancy engagements with Motmans Digital "
        "Vision, Rivr, and Stck.Pro have added experience in client coordination, delivery planning, and "
        "cross-functional collaboration.",
    )
    skill_line(
        doc,
        "IT Infrastructure & Operations",
        "Built a strong infrastructure foundation as a Server Administrator at Dhaka Web Host, working with "
        "Linux servers, DNS, email hosting, SSL, firewalls, and security troubleshooting. Apply this "
        "foundation to application deployment, VPS infrastructure, access control, backups, recovery, and "
        "production issue resolution.",
    )
    skill_line(
        doc,
        "Local & Global Engagement",
        "Experience across Bangladeshi organizations, including Dhaka Web Host, ADN Telecom PLC, and JS "
        "Group, alongside international work with Destinique Travel LLC (USA) and NybSys Inc. (USA). "
        "Additional consultancy and project engagements include Buzfi and European client projects "
        "associated with Motmans Digital Vision, Rivr, and Stck.Pro. Experienced in coordinating "
        "requirements and delivery across different clients and business environments. Professional "
        "communication in English and Bangla.",
    )
    skill_line(
        doc,
        "Achievements & Business Impact",
        "Built integrated business applications covering inventory, purchasing, sales, orders, payments, "
        "invoicing, and reporting. Delivered web applications, API integrations, dashboards, and automation "
        "to improve business workflows, and coordinated directly with corporate clients at ADN Telecom "
        "across requirements, deployment, and support.",
    )

    # ------------------------------------------------------------- Skills
    heading(doc, "IT Infrastructure and Security")
    skill_line(
        doc,
        "Servers and hosting",
        "Linux, DNS, email, SSL, firewall, Cloudflare, AWS (EC2, RDS, SES), DigitalOcean, Vultr, Docker, "
        "deployments, performance and uptime checks",
    )
    skill_line(
        doc,
        "Backup and recovery",
        "Database and server backups, and restoring service after a failure",
    )
    skill_line(
        doc,
        "Access and security",
        "User roles and permissions, account setup and removal, password and MFA support, patching, "
        "antivirus and device protection",
    )
    skill_line(
        doc,
        "Day-to-day IT",
        "User support, issue tracking in Jira and Trello, device and asset records, written documentation, "
        "and follow-up with hosting, cloud and other service providers",
    )

    heading(doc, "Business Applications and Digital Delivery")
    skill_line(
        doc,
        "Systems I have built or run",
        "ERP, POS, e-commerce, inventory, purchasing, sales, suppliers, customers, invoicing and billing, "
        "booking, property management, employee records",
    )
    skill_line(
        doc,
        "How I deliver",
        "Gather requirements from the people who will use the system, map their process, build, test with "
        "them, release, and support them after go-live. Connect systems through APIs and third-party services.",
    )
    skill_line(
        doc,
        "Reports",
        "Power BI and web dashboards, SQL reports for management, Google Analytics 4",
    )
    skill_line(
        doc,
        "Tools",
        "PHP, Laravel, JavaScript, Vue.js, React, Angular, MySQL, PostgreSQL, REST APIs, Git and GitHub",
    )

    # ------------------------------------------------------- Employment
    heading(doc, "Full-Time Employment")

    role(
        doc,
        "DESTINIQUE TRAVEL LLC",
        "Software Engineer  |  Florida, USA  |  Feb 2025 – Present",
        "Florida, USA-based luxury travel and vacation rental business.",
        [
            "Develop and maintain web applications and business automation solutions supporting travel inquiries, property quotations, and operational workflows.",
            "Build frontend features, backend integrations, and API-driven functionality, supporting deployments, troubleshooting, and ongoing improvements.",
        ],
        "Angular, PHP, Laravel, MySQL, JavaScript, REST APIs, Bootstrap, Cloudflare, Contabo VPS",
        projects=[
            ("Destinique Website (destinique.com)", "Travel and vacation rental website for property discovery and customer inquiries."),
            ("Quote Application (quote.destinique.com)", "Custom quotation platform supporting property quotes and travel booking workflows."),
            ("Destinique API Project", "Backend APIs and integrations supporting communication between the website, quote application, and business workflows."),
        ],
    )

    role(
        doc,
        "JS GROUP — JS ENTERPRISE / PAIMART",
        "Technical Manager, JS ERP  |  Dhaka, Bangladesh  |  Jan 2023 – Jan 2025",
        "Bangladesh-based FMCG importer and distributor with nationwide coverage, 1,000+ SKUs, 30,000+ general-trade outlets, and established modern-trade and e-commerce operations.",
        [
            "Promoted from Deputy Manager and independently designed and developed JS ERP, integrating sales, purchasing, inventory, supplier and customer management, invoicing, and billing into a unified business system.",
            "Managed the development lifecycle from requirements gathering and process mapping through implementation, release, user support, and system stabilization.",
            "Oversaw management reporting, user access, backups, and production issue resolution after deployment.",
        ],
    )

    role(
        doc,
        "JS GROUP — JS ENTERPRISE / PAIMART",
        "Deputy Manager, PaiMart  |  Dhaka, Bangladesh  |  Jan 2021 – Jan 2023",
        "",
        [
            "Managed and enhanced business applications supporting products, inventory, orders, customers, suppliers, payments, and billing.",
            "Collaborated with sales, warehouse, accounts, and field sales teams to digitize workflows and reduce manual processes.",
            "Supported user administration, access control, password and MFA issues, IT asset tracking, and device security.",
        ],
        "PHP, Laravel, MySQL, JavaScript, REST APIs, Git, GitHub, ERP/POS, VPS Hosting, Server Deployment, Vultr, AWS, DigitalOcean",
        projects=[
            ("JS ERP", "Integrated business operations and workflow management system."),
            ("PaiPOS", "ERP/POS platform for sales and business operations."),
            ("Operation (operation.paikaree.com.bd)", "Delivery management and order fulfillment system."),
            ("Paimart (paimart.painest.com)", "E-commerce platform."),
            ("Paibridge", "Multivendor e-commerce platform."),
            ("Paisoftdiaper.com", "UAE-focused brand website built with WordPress."),
        ],
    )

    role(
        doc,
        "ADN TELECOM PLC",
        "Assistant Manager, Engineering  |  Gulshan, Dhaka, Bangladesh  |  Oct 2019 – Jan 2021",
        "Bangladesh-based telecom, connectivity, and enterprise ICT solutions.",
        [
            "Coordinated with corporate clients, including Nitol Niloy (NTRAC), Goods in Motion (GIM), Nadia Furniture, Esquire Electronics, and FHDL, from requirements gathering through deployment and support, including vendor coordination.",
            "Contributed to in-house digital products, including RoboKet (ADN Diginet), ADN SMS, ADN Email, and ADN Server, supporting development and enhancements.",
            "Contributed to the NGTE Welfare Trust government project, supporting the delivery of a digital system for service and workflow management.",
        ],
        "PHP, Laravel, MySQL, JavaScript, REST APIs, Git, GitHub, Vultr VPS, Server Deployment, DNS Management, Firewall Configuration, Server Security",
        extras=[
            ("Key Clients & Projects", "Nitol Niloy (NTRAC) | Goods in Motion (GIM) | Nadia Furniture | Esquire Electronics | FHDL | NGTE Welfare Trust (Government Project)"),
            ("In-house Products", "RoboKet | ADN SMS | ADN Email | ADN Server"),
        ],
    )

    role(
        doc,
        "NYBSYS INC. (USA)",
        "Software Engineer  |  South Africa  |  Dec 2017 – Sep 2019",
        "USA-based technology, telecommunications, and digital infrastructure solutions.",
        [
            "Developed web frontends and API integrations while collaborating with distributed teams, supporting testing and releases.",
            "Contributed to CDAS, a government payroll-deductions platform in Eswatini, focusing on frontend development.",
            "Worked on Drinkwell web solutions and the SLAS student-loan application portal.",
        ],
        "PHP, JavaScript, HTML5, CSS3, API Integration, Git",
        projects=[
            ("CDAS (cdas.co.sz)", "Central Deductions Administration System for government payroll deductions in Eswatini."),
            ("SLAS (slas.gov.sz)", "Eswatini student-loan portal supporting applications, status tracking, and repayment information."),
            ("Drinkwell (drinkwell.com)", "Web development contributions while at NybSys."),
        ],
        tools_label="Technologies",
    )

    role(
        doc,
        "DHAKA WEB HOST LTD.",
        "Server Administrator  |  Dhaka, Bangladesh  |  2014 – 2017",
        "Bangladesh-based web hosting, domain, and server solutions provider.",
        [
            "Administered Linux servers supporting client hosting services, including DNS, email hosting, SSL certificates, firewall configuration, and server security.",
            "Diagnosed and resolved server, website, and email issues to maintain service availability and reliability.",
            "Supported client hosting environments and day-to-day infrastructure operations, including domain configuration, troubleshooting, and security maintenance.",
        ],
        "Linux, DNS Management, Email Hosting, SSL/TLS, Firewall Configuration, Server Security, Web Hosting",
        extras=[("Key Clients", "US-Bangla Airlines | Aftab Group | Bangla Academy | Dhaka Life Insurance")],
    )

    role(
        doc,
        "ICT DIVISION — LEARNING & EARNING DEVELOPMENT PROJECT",
        "Trainer  |  Dhaka, Bangladesh  |  Dec 2016 – Apr 2017",
        "Bangladesh government initiative focused on ICT skills and freelancing training.",
        [
            "Trained about 200 learners in HTML, CSS, PHP, WordPress and MySQL through hands-on lab sessions.",
        ],
    )

    # ------------------------------------------------------ Local engagements
    heading(doc, "Local Engagements")
    para(doc, "Direct client work, held alongside full-time employment.", size=SMALL, italic=True, color=MUTED, after=1)
    entry(doc, "Spiretec Limited", "Consultant  |  Bangladesh", "", "Direct consultancy for a local client, not through an employer.")
    entry(doc, "Sysba Trade Brand", "Local client engagement  |  Bangladesh", "", "Direct client work for a local business.")

    # ----------------------------------------------------- Global engagements
    heading(doc, "Global Engagements")
    para(
        doc,
        "Engineering management for clients in Belgium, Spain and Italy, and technology work for a US startup, "
        "held alongside full-time employment.",
        size=SMALL,
        italic=True,
        color=MUTED,
        after=1,
    )
    entry(
        doc,
        "BUZFI LLC",
        "Head of Technology  |  Jamaica, USA  |  Aug 2023 – Dec 2025",
        "Jamaica, New York, USA-based startup digital marketplace connecting buyers and sellers.",
        "Ran technology for the platform: a Next.js frontend on Vercel, a Laravel API, and Cloudflare with VPS hosting. "
        "Worked on performance, reliability, caching and security with web, API and mobile contributors.",
    )
    entry(
        doc,
        "Motmans Digital Vision",
        "Head of Development  |  Belgium  |  2020 – 2023",
        "Belgium-based IT consultancy specializing in business applications and automation.",
        "Managed engineering for product, web and business-application delivery for Belgian clients.",
    )
    entry(
        doc,
        "Rivr (rivr.be)",
        "Engineering Manager  |  Belgium  |  2019 – 2023",
        "Belgium-based digital brand specializing in web design, branding, and online marketing.",
        "Managed engineering for web delivery, including work connected with Orka Group, an SAP services firm. "
        "My part was engineering delivery.",
    )
    entry(
        doc,
        "Stck.Pro",
        "Technical Project Manager  |  Spain  |  2022 – 2023",
        "A stock market news platform built for users in the USA and the EU.",
        "Technical Project Manager for the platform build.",
    )

    # -------------------------------------------------------------- Education
    heading(doc, "Education")
    p = para(doc, "B.Sc. in Computer Science & Engineering", size=11, bold=True, before=1, after=0)
    keep_with_next(p)
    para(doc, "Dhaka International University  |  2010 – 2015", size=SMALL, color=MUTED)

    heading(doc, "Professional Training")
    p = para(doc, "Web Application Development with PHP", size=10.5, bold=True, before=1, after=0)
    keep_with_next(p)
    para(
        doc,
        "Bangladesh Association of Software and Information Services (BASIS)  |  2015",
        size=SMALL,
        color=MUTED,
    )
    p = para(doc, "Web Design & Development", size=10.5, bold=True, before=2, after=0)
    keep_with_next(p)
    para(doc, "Daffodil Institute of IT (DIIT)  |  2012", size=SMALL, color=MUTED)

    heading(doc, "Languages")
    para(doc, "Bangla (native)  |  English (professional working proficiency)")

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_v2.docx")
        doc.save(alt)
        print("locked; saved", alt)


if __name__ == "__main__":
    build()
