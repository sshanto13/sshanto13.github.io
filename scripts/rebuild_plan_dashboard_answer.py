"""Dashboard experience answer — max 500 characters."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn

OUT = Path(
    r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files"
    r"\Sumon_Ray_Plan_International_Dashboard_Experience_Answer.docx"
)

INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)
ACCENT = RGBColor(0x0F, 0x76, 0x6E)

ANSWER = (
    "I design web, Excel, and Power BI dashboards for operational "
    "monitoring. At Destinique, I support web dashboards for visitor "
    "tracking, promotions, booking/property records, and reporting "
    "(MySQL/SQL, query tuning, analytics). At PaiMart (4+ years), I built "
    "reports and dashboard views for inventory, orders, sales, payments, "
    "and ERP/POS data, focusing on accuracy and performance. I use Excel "
    "for cleaning, validation, formulas, pivots, and management reports, "
    "plus BI views for decision support."
)


def set_run(run, *, size=11, bold=False, color=INK, name="Calibri"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def para(doc, text="", *, size=11, bold=False, color=INK, before=0, after=8):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.15
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, color=color)


def build():
    n = len(ANSWER)
    if n > 500:
        raise SystemExit(f"Too long: {n} characters")

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    para(doc, "SUMON RAY", size=14, bold=True, after=2)
    para(
        doc,
        "Application Response — Plan International Bangladesh",
        size=10,
        bold=True,
        color=ACCENT,
        after=2,
    )
    para(
        doc,
        "Position: Data Management Specialist (B-STRONG / EGPP+)",
        size=10,
        color=MUTED,
        after=10,
    )
    para(doc, "Question", size=11, bold=True, after=4)
    para(
        doc,
        "Describe your experience designing and implementing cloud-based "
        "Power BI / web monitoring / Excel-based dashboards. (Maximum 500 characters)",
        size=10,
        color=MUTED,
        after=10,
    )
    para(doc, "Answer (copy into form)", size=11, bold=True, after=6)
    para(doc, ANSWER, size=11, after=10)
    para(doc, f"Character count: {n} / 500", size=9, color=MUTED, after=0)

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_500chars.docx")
        doc.save(alt)
        print("locked; saved", alt)
    print("chars", n)
    print(ANSWER)


if __name__ == "__main__":
    build()
