"""Improved cover letter for Plan International Data Management Specialist."""
from pathlib import Path
from datetime import date
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

OUT = Path(
    r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files"
    r"\Cover_Plan-international.docx"
)

INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)
ACCENT = RGBColor(0x0F, 0x76, 0x6E)


def set_run(run, *, size=11, bold=False, color=INK, name="Calibri"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def para(doc, text="", *, size=11, bold=False, color=INK, before=0, after=8, align="left", space=1.15):
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = space
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, color=color)
    return p


def build():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # Sender block
    para(doc, "SUMON RAY", size=16, bold=True, after=2)
    para(
        doc,
        "Data Management | MIS, Database Systems & Reporting",
        size=10,
        bold=True,
        color=ACCENT,
        after=2,
    )
    para(
        doc,
        "Dhaka, Bangladesh  |  sumancsediu@gmail.com  |  +880 1952 499680",
        size=10,
        color=MUTED,
        after=1,
    )
    para(
        doc,
        "LinkedIn: linkedin.com/in/sshanto13  |  Portfolio: sshanto13.github.io",
        size=10,
        color=MUTED,
        after=12,
    )

    # Date + recipient
    today = date.today().strftime("%d %B %Y")
    para(doc, today, size=11, after=10)
    para(doc, "Hiring Manager", size=11, bold=True, after=1)
    para(doc, "Plan International Bangladesh", size=11, after=1)
    para(doc, "Re: Data Management Specialist – B-STRONG / EGPP+", size=11, bold=True, after=12)

    para(doc, "Dear Hiring Manager,", size=11, after=10)

    para(
        doc,
        "I am writing to apply for the Data Management Specialist position with Plan International "
        "Bangladesh under the Bangladesh Sustainable Recovery, Emergency Preparedness and Response "
        "Project (B-STRONG). I am a Bangladeshi citizen with a B.Sc. in Computer Science & Engineering "
        "and more than eight years of professional experience building and maintaining database-driven "
        "information systems, operational reporting, and MIS-style platforms that keep records accurate, "
        "usable, and available for decision-making.",
        after=10,
    )

    para(
        doc,
        "The role’s focus on establishing and maintaining an effective data management and MIS function "
        "for EGPP+ — including systematic capture, verification, digitization, quality control, and "
        "reporting of beneficiary, sub-project, attendance, and payment-related data — closely matches "
        "how I have worked in production environments. Across my career I have designed and optimized "
        "relational databases (MySQL/PostgreSQL), written and tuned SQL, resolved data inconsistencies, "
        "supported operational teams with reliable reports and dashboards, and improved information "
        "flows from capture through validation to management reporting.",
        after=10,
    )

    para(
        doc,
        "A particularly relevant experience is my contribution to the NGTE Welfare Trust government "
        "information and application management system while at ADN Telecom PLC. That work involved "
        "structured application processing, tracking, document generation, and administrative workflows — "
        "where completeness, consistency, and reliable record handling mattered daily. Later, at PaiMart "
        "(JS Enterprise) and Destinique Travel LLC, I worked extensively with high-volume operational "
        "data across inventory, orders, payments, visitor tracking, booking/property records, and "
        "business reporting. I regularly investigated slow queries and data/integration issues, "
        "improved database performance, and delivered dashboards and custom reports that operations "
        "and management teams used for monitoring.",
        after=10,
    )

    para(
        doc,
        "Beyond technical delivery, I am comfortable working with non-technical users, clarifying "
        "requirements, and guiding colleagues through data and system processes. As a trainer under "
        "an ICT Division programme, I facilitated practical sessions for around 200 learners — "
        "experience I can apply when supporting Data Entry Operators and field teams on MIS procedures, "
        "data quality checks, and discrepancy resolution. I also bring preferred strengths noted in the "
        "vacancy: web/software development, ERP-style business systems, Power BI and dashboard work, "
        "and strong day-to-day use of Microsoft Excel for cleaning, validation, formulas, and reporting.",
        after=10,
    )

    para(
        doc,
        "I am motivated by Plan International’s mission and by the practical impact of trustworthy "
        "project data on flood recovery, livelihoods, and community resilience. I am based in Dhaka "
        "and available for the travel required across Feni, Brahmanbaria, Chattogram, and related "
        "Upazilas. I would welcome the opportunity to contribute reliable data management, MIS "
        "discipline, and clear reporting support to the Team Leader, PIU, and field implementation teams.",
        after=10,
    )

    para(
        doc,
        "Thank you for considering my application. I have attached my CV and would be glad to discuss "
        "how my experience can support EGPP+ data quality, monitoring, and evidence-based decision-making.",
        after=14,
    )

    para(doc, "Sincerely,", size=11, after=28)
    para(doc, "Sumon Ray", size=11, bold=True, after=2)
    para(doc, "sumancsediu@gmail.com  |  +880 1952 499680", size=10, color=MUTED, after=0)

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name("Cover_Plan-international_improved.docx")
        doc.save(alt)
        print("locked; saved", alt)


if __name__ == "__main__":
    build()
