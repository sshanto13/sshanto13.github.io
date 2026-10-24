"""Tailored CV for Plan International Bangladesh — Data Management Specialist."""
from pathlib import Path
import shutil
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = Path(
    r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files"
    r"\Sumon_Ray_Data_Management_Plan_International_2026.docx"
)
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


def set_para_spacing(p, before=0, after=4, line=1.12):
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


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        tcBorders.append(element)
    tcPr.append(tcBorders)


def heading(doc, text):
    p = doc.add_paragraph()
    set_para_spacing(p, before=9, after=3, line=1.05)
    r = p.add_run(text.upper())
    set_run(r, size=10.5, bold=True, color=ACCENT)
    add_bottom_border(p)


def body(doc, text, *, size=9.5, bold=False, before=0, after=2):
    p = doc.add_paragraph()
    set_para_spacing(p, before=before, after=after, line=1.12)
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=INK)
    return p


def muted(doc, text, *, size=9, before=0, after=2):
    p = doc.add_paragraph()
    set_para_spacing(p, before=before, after=after, line=1.08)
    r = p.add_run(text)
    set_run(r, size=size, bold=False, color=MUTED)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    set_para_spacing(p, before=0, after=1.5, line=1.1)
    if p.runs:
        p.runs[0].text = text
        set_run(p.runs[0], size=9.2, color=INK)
    else:
        r = p.add_run(text)
        set_run(r, size=9.2, color=INK)


def skill_line(doc, label, items):
    p = doc.add_paragraph()
    set_para_spacing(p, before=1, after=1.5, line=1.1)
    r1 = p.add_run(f"{label}: ")
    set_run(r1, size=9.2, bold=True, color=INK)
    r2 = p.add_run(items)
    set_run(r2, size=9.2, bold=False, color=INK)


def build():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.48)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Header: text left, photo right
    header = doc.add_table(rows=1, cols=2)
    text_cell = header.cell(0, 0)
    photo_cell = header.cell(0, 1)
    set_cell_border(text_cell)
    set_cell_border(photo_cell)
    text_cell.width = Inches(5.55)
    photo_cell.width = Inches(1.25)

    p_name = text_cell.paragraphs[0]
    set_para_spacing(p_name, before=0, after=1, line=1.0)
    set_run(p_name.add_run("SUMON RAY"), size=17, bold=True, color=INK)

    p_title = text_cell.add_paragraph()
    set_para_spacing(p_title, before=0, after=2, line=1.05)
    set_run(
        p_title.add_run("Data Management Specialist | MIS, Database & Reporting"),
        size=10.5,
        bold=True,
        color=ACCENT,
    )

    p_tag = text_cell.add_paragraph()
    set_para_spacing(p_tag, before=0, after=3, line=1.08)
    set_run(
        p_tag.add_run(
            "Building and maintaining data systems, MIS workflows, databases, dashboards, "
            "and reliable reporting for operational decision-making."
        ),
        size=9,
        color=MUTED,
    )

    p_contact = text_cell.add_paragraph()
    set_para_spacing(p_contact, before=0, after=1, line=1.05)
    set_run(
        p_contact.add_run(
            "Dhaka, Bangladesh  |  sumancsediu@gmail.com  |  +880 1952 499680  |  Bangladeshi Citizen"
        ),
        size=8.5,
        color=INK,
    )

    p_links = text_cell.add_paragraph()
    set_para_spacing(p_links, before=0, after=1, line=1.05)
    set_run(
        p_links.add_run(
            "LinkedIn: linkedin.com/in/sshanto13  |  Portfolio: sshanto13.github.io  |  "
            "GitHub: github.com/sshanto13"
        ),
        size=8.5,
        color=MUTED,
    )

    p_photo = photo_cell.paragraphs[0]
    set_para_spacing(p_photo, before=0, after=0, line=1.0)
    p_photo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if PHOTO.exists():
        p_photo.add_run().add_picture(str(PHOTO), width=Inches(1.1))

    spacer = doc.add_paragraph()
    set_para_spacing(spacer, before=1, after=1, line=1.0)

    # Profile — mapped to JD without fabricating WB/IBAS++
    heading(doc, "Professional Profile")
    body(
        doc,
        "Computer Science graduate and software professional with 8+ years of hands-on experience "
        "in data management, database systems, MIS-style operational platforms, reporting, and "
        "information workflows across e-commerce, enterprise, government ICT, and travel/property "
        "operations. Strong track record designing, maintaining, and improving relational databases; "
        "validating and structuring operational records; building dashboards and reports; and "
        "supporting teams that depend on accurate, timely data.",
        after=2,
    )
    body(
        doc,
        "Proven ability in database design, SQL query writing, performance tuning and optimization, "
        "data quality troubleshooting, and end-to-end information flows from capture and digitization "
        "through verification, storage, reporting, and decision support. Comfortable coordinating with "
        "business/operations users, guiding colleagues on data processes, resolving discrepancies, and "
        "delivering clear summaries for monitoring and management use. Available for Dhaka-based work "
        "with regular field travel as required.",
        after=2,
    )
    body(
        doc,
        "Core technical strengths include MySQL/PostgreSQL, SQL, Microsoft Excel (formulas, cleaning, "
        "validation, pivot-style analysis, and reporting), Power BI / operational dashboards, PHP/Laravel "
        "and related web systems, REST APIs, ERP/business platforms, and production systems support. "
        "Eager to apply this background to EGPP+ data management and MIS responsibilities, including "
        "beneficiary, attendance, sub-project, and payment-related information.",
        after=3,
    )

    # Competencies — JD order
    heading(doc, "Core Competencies (Aligned to Role)")
    skill_line(
        doc,
        "Data Management & Quality",
        "Structured data capture, verification, digitization support, data validation, consistency checks, "
        "discrepancy resolution, record tracking, completeness/accuracy focus",
    )
    skill_line(
        doc,
        "MIS & Operational Reporting",
        "MIS-style business systems, operational dashboards, custom reports, monitoring summaries, "
        "management reporting, Google Analytics / tracking integrations",
    )
    skill_line(
        doc,
        "Database Management",
        "MySQL, PostgreSQL, SQL, schema design, indexing awareness, query optimization, "
        "database performance tuning, troubleshooting, MongoDB/Redis where relevant",
    )
    skill_line(
        doc,
        "Microsoft Excel & Analysis",
        "Formulas, data cleaning, validation, filtering, lookup-style analysis, pivot tables/summaries, "
        "spreadsheet reporting for operational and management use",
    )
    skill_line(
        doc,
        "Dashboards & BI",
        "Power BI, custom dashboards, KPI-style operational views, report preparation for decision-makers",
    )
    skill_line(
        doc,
        "Systems Advantage (Preferred)",
        "Web development, software development, Cloud ERP / business systems (e-commerce, inventory, "
        "billing, booking), REST APIs, process automation, VPS/cloud deployment support",
    )
    skill_line(
        doc,
        "Collaboration & Capacity Support",
        "Stakeholder requirements gathering, guidance for non-technical users, mentoring/training "
        "(including ICT Division training for ~200 learners), cross-team coordination, production support",
    )

    # Experience — reframed to data/MIS language
    heading(doc, "Professional Experience")

    body(doc, "DESTINIQUE TRAVEL LLC", bold=True, before=3, after=0)
    muted(doc, "Software Engineer  |  Dhaka, Bangladesh  |  Feb 2025 - Present")
    body(
        doc,
        "Own and improve data-driven operational systems supporting visitor tracking, promotions, "
        "property/booking records, dashboards, and day-to-day business reporting.",
        size=9,
        after=1,
    )
    for b in [
        "Maintain and enhance systems that capture, store, and surface operational records used by business teams for monitoring and reporting.",
        "Build and support dashboards and REST API-backed views so teams can review activity, track progress, and act on timely information.",
        "Work with MySQL structures and SQL queries; investigate slow queries and data/integration issues affecting completeness, accuracy, and usability.",
        "Support analytics and tracking integrations; help keep operational data flows reliable from capture through reporting.",
        "Translate stakeholder requirements into practical data/process solutions; assist users when records, reports, or workflows need clarification or fixes.",
        "Contribute to application reliability, deployment, and production maintenance so information systems remain available for field/office use.",
    ]:
        bullet(doc, b)
    muted(doc, "Tools: PHP, Laravel, MySQL, SQL, Angular, JavaScript, REST API, dashboards, Cloudflare/VPS", after=3)

    body(doc, "PAIMART BY PAINEST - JS ENTERPRISE LTD.", bold=True, before=1, after=0)
    muted(doc, "Assistant Manager, Engineering  |  Dhaka, Bangladesh  |  Jan 2021 - Jan 2025")
    body(
        doc,
        "Led and delivered data-heavy business platforms across e-commerce, distribution, inventory, "
        "orders, payments, ERP/POS, and operational reporting for 4+ years.",
        size=9,
        after=1,
    )
    for b in [
        "Designed and maintained relational database structures and SQL-driven workflows for products, customers, suppliers, inventory, orders, sales, and distribution data.",
        "Supported high-volume operational data used in back-office processes; improved query performance and resolved data-related production issues as volume grew.",
        "Developed reports, dashboards, and MIS-style operational views to help teams monitor sales, inventory, billing, and day-to-day business status.",
        "Worked across end-to-end information flows: capture through business applications, validation/troubleshooting of inconsistencies, and usable reporting outputs.",
        "Exposed APIs and integrated systems so storefront/mobile and back-office teams worked from consistent shared data.",
        "Collaborated with business stakeholders to clarify data needs, prioritize fixes, and keep records and processes usable for operations staff.",
    ]:
        bullet(doc, b)
    muted(doc, "Tools: PHP, Laravel, MySQL, SQL, JavaScript, REST API, ERP/POS, reporting/dashboards, Git, VPS", after=3)

    body(doc, "ADN TELECOM PLC", bold=True, before=1, after=0)
    muted(doc, "Assistant Manager, Engineering  |  Gulshan, Dhaka, Bangladesh  |  Oct 2019 - Jan 2021")
    body(
        doc,
        "Built enterprise and government ICT information systems involving structured application "
        "records, tracking, document generation, and administrative workflows.",
        size=9,
        after=1,
    )
    for b in [
        "Contributed to NGTE Welfare Trust government information and application management system supporting application processing, tracking, documents, and admin workflows.",
        "Developed Laravel/MySQL backends and APIs for storing, retrieving, and managing structured administrative and application data.",
        "Supported schema design, SQL work, testing, deployment, and resolution of application/database issues affecting record accuracy and process flow.",
        "Worked with teams to keep information workflows practical for administrative users handling processing and tracking tasks.",
    ]:
        bullet(doc, b)
    muted(doc, "Tools: PHP, Laravel, MySQL, SQL, JavaScript, REST API, Bootstrap", after=3)

    body(doc, "NYBSYS INC.", bold=True, before=1, after=0)
    muted(doc, "Software Engineer  |  Mirpur, Dhaka, Bangladesh  |  Dec 2017 - Sep 2019")
    for b in [
        "Developed web interfaces and API integrations for applications used by distributed teams.",
        "Converted requirements and tickets into working screens and data interactions; supported testing, debugging, and delivery.",
    ]:
        bullet(doc, b)
    muted(doc, "Tools: PHP, JavaScript, HTML, CSS, API Integration", after=3)

    body(doc, "ICT DIVISION - LEARNING & EARNING DEVELOPMENT PROJECT", bold=True, before=1, after=0)
    muted(doc, "Trainer  |  Dhaka, Bangladesh  |  Dec 2016 - Apr 2017")
    for b in [
        "Trained approximately 200 learners in HTML, CSS, PHP, WordPress, and MySQL through practical lab sessions.",
        "Built experience explaining technical concepts clearly - relevant to guiding data entry and field teams on MIS processes.",
    ]:
        bullet(doc, b)

    # Projects — role-relevant first
    heading(doc, "Selected Relevant Projects")
    projects = [
        (
            "NGTE Welfare Trust - Government Information System",
            "https://apps.ngte-welfaretrust.gov.bd/",
            "Laravel | PHP | MySQL | JavaScript",
            "Government application and information management platform supporting application processing, tracking, document generation, and administrative record workflows.",
        ),
        (
            "Destinique Business Dashboard",
            "https://quote.destinique.com/destin/dashboard/",
            "PHP | MySQL | JavaScript | REST API | Cloudflare",
            "Internal operational platform for visitor tracking, promotions, property/booking records, reporting, and day-to-day business monitoring.",
        ),
        (
            "E-commerce Operations - Back Office",
            "https://operation.paikaree.com.bd/",
            "Laravel | PHP | MySQL | JavaScript",
            "Back-office system for product, order, and inventory operations data used by internal teams.",
        ),
        (
            "JS ERP - Sales, Billing & Inventory",
            "https://js-erp.paitechno.com/",
            "Laravel | PHP | MySQL | JavaScript",
            "Business management system covering sales, invoicing, billing, inventory, and operational processes with structured transactional data.",
        ),
        (
            "PaiMart - Wholesale E-commerce Platform",
            "https://paimart.painest.com/",
            "Laravel | PHP | MySQL | JavaScript | REST API",
            "Wholesale e-commerce and operations platform spanning products, inventory, orders, payments, ERP/POS, reporting, and APIs.",
        ),
    ]
    for title, url, stack, desc in projects:
        p = doc.add_paragraph()
        set_para_spacing(p, before=2, after=0, line=1.05)
        set_run(p.add_run(title), size=9.5, bold=True, color=INK)
        set_run(p.add_run(f"  ({url})"), size=8, color=MUTED)
        muted(doc, stack, size=8, after=0)
        body(doc, desc, size=9, after=1)

    # Role fit note — honest about preferred items
    heading(doc, "Role Fit Summary")
    for b in [
        "Meets education requirement: B.Sc. in Computer Science & Engineering (4-year honours).",
        "Exceeds minimum experience: 8+ years involving database management, MIS-style systems, monitoring/reporting, and operational data workflows (requirement: 2+ years).",
        "Strong match on database performance tuning/optimization, SQL, dashboards, web/software development, and Cloud ERP/business systems experience.",
        "Practical Excel and reporting experience supporting operational analysis; Power BI and dashboard delivery as an advantage area.",
        "Government ICT information-system experience (NGTE Welfare Trust) involving structured application records and tracking workflows.",
        "No prior hands-on use of IBAS++ or World Bank Client Connection listed; ready to learn project-specific platforms quickly and apply established data-quality practices.",
        "Willing and available for regular travel to Feni, Brahmanbaria, Chattogram and related Upazilas as required.",
    ]:
        bullet(doc, b)

    heading(doc, "Education")
    body(doc, "B.Sc. in Computer Science & Engineering", bold=True, before=1, after=0)
    muted(doc, "Dhaka International University  |  2010 - 2015")

    heading(doc, "Professional Training")
    body(doc, "Web Application Development with PHP", bold=True, before=1, after=0)
    muted(doc, "BASIS / BITM  |  2015")
    body(doc, "Web Design & Development", bold=True, before=2, after=0)
    muted(doc, "Daffodil Institute of IT (DIIT)  |  2012")

    # Save
    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_improved.docx")
        doc.save(alt)
        print("locked; saved", alt)
    print("photo", PHOTO.exists())


if __name__ == "__main__":
    build()
