"""Parse the updated CV (Sumon_Ray_CV_ACI_Lead_IT_Digitalisation_2026.pdf.docx) into the same plain
content model that cv_model.parse() returns. Text is copied verbatim; nothing is reworded here."""
import re
import docx
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HEADINGS = {
    "PROFESSIONAL SUMMARY", "TECHNICAL EXPERIENCE", "PROFESSIONAL EXPERIENCE", "GLOBAL PROJECT ENGAGEMENTS",
    "GOVERNMENT INITIATIVES & TRAINING", "EDUCATION", "TRAINING & PROFESSIONAL CERTIFICATIONS",
}


def ptxt(p):
    out = []
    for n in p.iter():
        if n.tag == W + "t":
            out.append(n.text or "")
        elif n.tag == W + "tab":
            out.append("\t")
    return "".join(out)


def first_bold(p):
    r = p.find(qn("w:r"))
    if r is None:
        return False
    rp = r.find(qn("w:rPr"))
    return rp is not None and rp.find(qn("w:b")) is not None


def style(p):
    pPr = p.find(qn("w:pPr"))
    st = pPr.find(qn("w:pStyle")) if pPr is not None else None
    return st.get(qn("w:val")) if st is not None else ""


def parse(path):
    d = docx.Document(str(path))
    kids = [e for e in d.element.body.iterchildren() if e.tag in (W + "p", W + "tbl")]
    m = dict(name="", headline="", contacts=[], summary="", highlights=[], tech_cards=[], experience=[],
             consultancy=[], gov=[], education=[], training=[], headings={})
    head = [ptxt(p).strip() for p in list(kids[0].iter(qn("w:tc")))[0].findall(qn("w:p"))]
    m["name"], m["headline"] = head[0], head[1]

    mode, cur, card, last_title = None, None, None, None
    for e in kids[1:]:
        if e.tag == W + "tbl":
            if mode == "PROFESSIONAL SUMMARY":
                for c in e.iter(qn("w:tc")):
                    t = [ptxt(p).strip() for p in c.findall(qn("w:p")) if ptxt(p).strip()]
                    if t:
                        m["highlights"].append((t[0], t[1] if len(t) > 1 else ""))
            continue
        t = ptxt(e).replace("\xa0", " ").strip()
        if not t:
            continue
        if t.upper() in HEADINGS:
            mode = t.upper()
            m["headings"][mode] = t
            cur = card = None
            continue
        if mode == "PROFESSIONAL SUMMARY":
            m["summary"] = t
        elif mode == "TECHNICAL EXPERIENCE":
            if re.match(r"^\d+\.\s", t):
                card = dict(title=t, lines=[])
                m["tech_cards"].append(card)
            else:
                label, _, rest = t.partition(":")
                card["lines"].append((label.strip() + ":", rest.strip()))
        elif mode in ("PROFESSIONAL EXPERIENCE", "GLOBAL PROJECT ENGAGEMENTS", "GOVERNMENT INITIATIVES & TRAINING"):
            full = ptxt(e).replace("\xa0", " ")
            lst = {"PROFESSIONAL EXPERIENCE": m["experience"], "GLOBAL PROJECT ENGAGEMENTS": m["consultancy"],
                   "GOVERNMENT INITIATIVES & TRAINING": m["gov"]}[mode]
            if style(e) == "ListBullet":
                label, sep, rest = t.partition(":")
                if first_bold(e) and sep:
                    cur["bullets"].append(("label", label.strip() + ":", rest.strip()))
                else:
                    cur["bullets"].append(("plain", t))
            elif t.startswith("Technologies:"):
                cur["tech"] = t[len("Technologies:"):].strip()
            elif "\t" in full:
                parts = [x.strip() for x in full.split("\t") if x.strip()]
                cur = dict(title=parts[0], date=parts[1], company=None, intro=None, bullets=[], tech=None)
                lst.append(cur)
            elif first_bold(e) and cur["company"] is None and mode != "GOVERNMENT INITIATIVES & TRAINING":
                cur["company"] = t
            else:
                cur["intro"] = t
        elif mode == "EDUCATION":
            if "\t" in ptxt(e):
                parts = [x.strip() for x in ptxt(e).split("\t") if x.strip()]
                m["education"].append(dict(title=parts[0], date=parts[1] if len(parts) > 1 else "", school=""))
            else:
                m["education"][-1]["school"] = t
        elif mode == "TRAINING & PROFESSIONAL CERTIFICATIONS":
            if first_bold(e):
                m["training"].append([t, ""])
            else:
                m["training"][-1][1] = t
    m["training"] = [tuple(x) for x in m["training"]]
    return m
