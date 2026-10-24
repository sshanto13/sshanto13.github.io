"""Take the user's v3 ACI Godrej CV (format they like) and replace the JS Group
block with two clear roles: Technical Manager (JS ERP) and Deputy Manager (PaiMart).
All other formatting and text in v3 is left untouched."""
import copy
from pathlib import Path
from docx import Document

ROOT = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
SRC = ROOT / "Sumon_Ray_CV_Lead_IT_Digitalization_v3.docx"
OUT = ROOT / "Sumon_Ray_CV_Lead_IT_Digitalization_v4.docx"

doc = Document(SRC)
paras = doc.paragraphs

# locate JS block: starts at paragraph whose text contains "JS GROUP", ends before ADN title
start = next(i for i, p in enumerate(paras) if "JS GROUP" in p.text and p.runs and p.runs[0].bold)
end = next(i for i, p in enumerate(paras) if p.text.startswith("Assistant Manager, Engineering"))
anchor = paras[end]._p  # insert new content before this element

# templates (deep copies of v3 elements so formatting is identical)
t_title = copy.deepcopy(paras[end]._p)          # bold navy "Title | COMPANY"
t_dates = copy.deepcopy(paras[end + 1]._p)      # italic dates line
t_about = copy.deepcopy(paras[start + 2]._p)    # italic company description
t_bullet = copy.deepcopy(paras[start + 3]._p)   # plain bullet
t_label = copy.deepcopy(paras[start + 6]._p)    # bold label + normal text bullet
about_text = paras[start + 2].text
for i in range(start, end):
    paras[i]._p.getparent().remove(paras[i]._p)


def add(template, *texts):
    el = copy.deepcopy(template)
    anchor.addprevious(el)
    from docx.text.paragraph import Paragraph
    p = Paragraph(el, paras[0]._parent)
    for run, text in zip(p.runs, texts):
        run.text = text
    for run in p.runs[len(texts):]:
        run._r.getparent().remove(run._r)
    return p


COMPANY = "|  JS GROUP — JS ENTERPRISE / PAIMART"

# ---- Technical Manager
add(t_title, "Technical Manager, JS ERP ", COMPANY)
add(t_dates, "Jan 2023 – Jan 2025  |  Dhaka, Bangladesh  |  Promoted from Deputy Manager")
add(t_about, about_text)
add(t_bullet, "Promoted from Deputy Manager and independently designed and developed JS ERP, integrating sales, purchasing, inventory, supplier and customer management, invoicing, and billing into a unified business system.")
add(t_bullet, "Managed the development lifecycle from requirements gathering and process mapping through implementation, release, user support, and system stabilization.")
add(t_bullet, "Oversaw management reporting, user access, backups, and production issue resolution after deployment.")

# ---- Deputy Manager
add(t_title, "Deputy Manager, PaiMart ", COMPANY)
add(t_dates, "Jan 2021 – Jan 2023  |  Dhaka, Bangladesh")
add(t_bullet, "Managed and enhanced business applications supporting products, inventory, orders, customers, suppliers, payments, and billing.")
add(t_bullet, "Collaborated with sales, warehouse, accounts, and field sales teams to digitize workflows and reduce manual processes.")
add(t_bullet, "Supported user administration, access control, password and MFA issues, IT asset tracking, and device security.")
add(t_label, "Key Projects & Platforms:",
    " JS ERP (integrated business operations and workflow system) | PaiPOS (ERP/POS for sales and operations) | "
    "Operation, operation.paikaree.com.bd (delivery management and order fulfillment) | "
    "Paimart, paimart.painest.com (e-commerce) | Paibridge (multivendor e-commerce) | "
    "Paisoftdiaper.com (UAE-focused brand website, WordPress).")
add(t_label, "Technologies & Infrastructure:",
    " PHP, Laravel, MySQL, JavaScript, REST APIs, Git, GitHub, ERP/POS, VPS Hosting, Server Deployment, Vultr, AWS, DigitalOcean.")

# keep section headings, role titles, date lines and company lines with what follows
for p in doc.paragraphs:
    if p.style.name == "Normal" and p.runs and (p.runs[0].bold or p.runs[0].italic):
        p.paragraph_format.keep_with_next = True

try:
    doc.save(OUT)
    print("saved", OUT)
except PermissionError:
    alt = OUT.with_name(OUT.stem + "_b.docx")
    doc.save(alt)
    print("locked; saved", alt)
