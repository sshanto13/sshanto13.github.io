"""One-page cover letter for The Asia Foundation Bangladesh — IT Officer."""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io")
OUT = ROOT / "public" / "files" / "Sumon_Ray_Cover_Letter_Asia_Foundation_2026.docx"

ACCENT = RGBColor(0x0F, 0x76, 0x6E)
INK = RGBColor(0x12, 0x15, 0x1A)
MUTED = RGBColor(0x4B, 0x55, 0x63)
SIZE = 10.5


def set_run(run, *, size=SIZE, bold=False, color=INK):
    run.font.name = "Calibri"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def para(doc, text="", *, size=SIZE, bold=False, color=INK, after=6, line=1.12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    if text:
        set_run(p.add_run(text), size=size, bold=bold, color=color)
    return p


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "12"), ("w:space", "4"), ("w:color", "0F766E")):
        bottom.set(qn(k), v)
    pBdr.append(bottom)
    pPr.append(pBdr)


def build():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.85)
        s.right_margin = Inches(0.85)

    para(doc, "SUMON RAY", size=17, bold=True, after=0, line=1.0)
    para(doc, "IT Systems & Support Professional", size=11, bold=True, color=ACCENT, after=2)
    p = para(
        doc,
        "Dhaka, Bangladesh  |  sumancsediu@gmail.com  |  +880 1952 499680  |  linkedin.com/in/sshanto13",
        size=9.5,
        color=MUTED,
        after=10,
    )
    bottom_border(p)

    para(doc, "28 September 2026", after=8)
    para(doc, "Hiring Committee", after=0)
    para(doc, "The Asia Foundation, Bangladesh", after=0)
    para(doc, "Dhaka", after=8)
    para(doc, "Subject: Application for the position of IT Officer", bold=True, after=8)

    para(doc, "Dear Hiring Committee,", after=6)

    for text in [
        "I am applying for the position of IT Officer at The Asia Foundation's Bangladesh office. "
        "I hold a B.Sc. in Computer Science and Engineering and bring more than 8 years of experience "
        "supporting business-critical systems, infrastructure, and users across ICT, e-commerce, "
        "government ICT projects, and an international company.",

        "My work has always been close to the people who depend on the systems. I provide Level 1–2 "
        "support: diagnosing issues, guiding non-technical staff, tracking incidents to resolution, "
        "and escalating to vendors or specialists when needed. This includes Windows 10/11 "
        "installation and troubleshooting, Microsoft 365 user support, basic LAN/Wi-Fi issues, "
        "printers and peripherals, and remote support.",

        "At Destinique Travel LLC, I manage Cloudflare and Linux VPS environments, including DNS, "
        "SSL/TLS, security settings, updates, and deployments, for a team working across locations. "
        "As Assistant Manager, Engineering at PaiMart by PaiNest, I supported ERP, POS, and "
        "e-commerce systems used by sales, inventory, billing, and back-office staff, managed user "
        "roles and access permissions, and handled production incidents and releases under "
        "controlled procedures. Earlier, I worked on government ICT projects and web hosting "
        "support covering DNS, email hosting, and SSL.",

        "The induction and cybersecurity-awareness parts of this role suit me well. I trained "
        "approximately 200 learners through a Government of Bangladesh ICT Division programme and "
        "regularly explain technical issues in plain language. I also maintain technical "
        "documentation and prepare dashboards and periodic reports for management.",

        "Although my job titles have been in engineering, much of my day-to-day work has been IT "
        "support and operations, and this is where I want to focus my career. I am drawn to The "
        "Asia Foundation because reliable, secure IT directly supports its development programmes "
        "in Bangladesh. I am comfortable following established IT, security, and compliance "
        "policies, working with regional and global IT teams, and providing after-hours support "
        "when required, and I am keen to deepen my skills in endpoint management tools such as "
        "Microsoft Intune.",

        "Thank you for considering my application. I would welcome the opportunity to discuss how "
        "I can support the Foundation's Bangladesh office.",
    ]:
        para(doc, text, after=7)

    para(doc, "Sincerely,", after=14)
    para(doc, "Sumon Ray", bold=True, after=0)

    try:
        doc.save(OUT)
        print("saved", OUT)
    except PermissionError:
        alt = OUT.with_name(OUT.stem + "_v2.docx")
        doc.save(alt)
        print("locked; saved", alt)


if __name__ == "__main__":
    build()
