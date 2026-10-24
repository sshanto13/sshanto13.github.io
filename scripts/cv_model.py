"""Parse the user's final CV (Sumon_Ray_CV_Reviewed_typography.docx) into a plain content model.
No text is altered here - every layout renders exactly this model."""
import sys
from pathlib import Path
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
SRC = BASE / "Sumon_Ray_CV_Lead_IT_Digitalisation_2026.pdf.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

HEADINGS = {
    "PROFESSIONAL SUMMARY", "TECHNICAL EXPERIENCE", "PROFESSIONAL EXPERIENCE", "GLOBAL PROJECT ENGAGEMENTS",
    "GOVERNMENT INITIATIVES & TRAINING", "EDUCATION", "TRAINING & PROFESSIONAL CERTIFICATIONS",
}


def ptxt(p_el):
    out = []
    for node in p_el.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == W + "tab" and node.getparent().tag == W + "r":
            out.append("\t")
    return "".join(out)


def cell_paras(tc):
    return [p for p in tc.findall(qn("w:p"))]


def runs_info(p_el):
    res = []
    for r in p_el.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        bold = rPr is not None and rPr.find(qn("w:b")) is not None
        font = None
        if rPr is not None and rPr.find(qn("w:rFonts")) is not None:
            font = rPr.find(qn("w:rFonts")).get(qn("w:ascii"))
        t = "".join((x.text or "") for x in r.findall(qn("w:t")))
        res.append((t, bold, font))
    return res


def parse(path=SRC):
    if str(path).endswith(".pdf.docx"):
        import cv_model_pdfdocx
        return cv_model_pdfdocx.parse(path)
    d = docx.Document(str(path))
    body = d.element.body
    kids = [e for e in body.iterchildren() if e.tag in (W + "p", W + "tbl")]
    m = dict(name="", headline="", contacts=[], summary="", highlights=[], tech_cards=[], experience=[],
             consultancy=[], gov=[], education=[], training=[], headings={})

    # --- hyperlinks of the header table
    rels = d.part.rels
    hl = {}
    for h in kids[0].iter(qn("w:hyperlink")):
        rid = h.get(qn("r:id"))
        txt = "".join(t.text or "" for t in h.iter(qn("w:t"))).strip()
        hl[txt] = rels[rid].target_ref
    tcs = list(kids[0].iter(qn("w:tc")))
    left = cell_paras(tcs[0])
    m["name"] = ptxt(left[0]).strip()
    m["headline"] = ptxt(left[1]).strip()
    for p in left[2:]:
        # contact paragraph: runs alternate emoji / text / tab-sep
        for r in runs_info(p):
            pass
        raw = []
        for r in p.iter(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            font = rPr.find(qn("w:rFonts")).get(qn("w:ascii")) if rPr is not None and rPr.find(qn("w:rFonts")) is not None else None
            t = "".join((x.text or "") for x in r.findall(qn("w:t"))).strip()
            if t and font != "Segoe UI Emoji":
                raw.append(t)
        for t in raw:
            m["contacts"].append((t, hl.get(t)))
    # tolerate contacts being in other cells
    mode = None
    section = None
    i = 1
    cur = None
    while i < len(kids):
        e = kids[i]
        i += 1
        if e.tag == W + "p":
            t = ptxt(e).strip()
            if t.upper() in HEADINGS and t:
                mode = t.upper()
                m["headings"][mode] = t
                cur = None
                continue
            if not t:
                continue
            if mode in ("PROFESSIONAL EXPERIENCE", "GLOBAL PROJECT ENGAGEMENTS", "GOVERNMENT INITIATIVES & TRAINING"):
                pPr = e.find(qn("w:pPr"))
                st = pPr.find(qn("w:pStyle")) if pPr is not None else None
                style = st.get(qn("w:val")) if st is not None else None
                full = ptxt(e)
                ri = runs_info(e)
                if style == "ListBullet":
                    # label = leading bold run
                    if ri and ri[0][1] and ri[0][0].strip().endswith(":"):
                        cur["bullets"].append(("label", ri[0][0].strip(), "".join(r[0] for r in ri[1:]).strip()))
                    else:
                        cur["bullets"].append(("plain", t))
                    continue
                if t.replace("\xa0", " ").strip().startswith("Technologies:"):
                    s = t.replace("\xa0", " ").strip()
                    cur["tech"] = s[len("Technologies:"):].strip()
                    continue
                if "\t" in full:
                    title, date = full.split("\t", 1)
                    cur = dict(title=title.strip(), date=date.strip(), company=None, intro=None, bullets=[], tech=None)
                    lst = {"PROFESSIONAL EXPERIENCE": m["experience"], "GLOBAL PROJECT ENGAGEMENTS": m["consultancy"],
                              "GOVERNMENT INITIATIVES & TRAINING": m["gov"]}[mode]
                    lst.append(cur)
                    continue
                # company line: Arial, bold, no italic. intro: italic
                italic = any(r.find(qn("w:rPr")) is not None and r.find(qn("w:rPr")).find(qn("w:i")) is not None
                             for r in e.iter(qn("w:r")))
                if italic:
                    cur["intro"] = t
                elif cur["company"] is None and mode != "GOVERNMENT INITIATIVES & TRAINING":
                    cur["company"] = t
                else:
                    cur["intro"] = t
        else:  # table
            tbl = e
            cells = list(tbl.iter(qn("w:tc")))
            if mode == "PROFESSIONAL SUMMARY":
                texts = [[ptxt(p).strip() for p in cell_paras(c) if ptxt(p).strip()] for c in cells]
                texts = [t for t in texts if t]
                if len(texts) == 1 and len(texts[0]) == 1:
                    m["summary"] = texts[0][0]
                else:
                    for t in texts:
                        m["highlights"].append((t[0], t[1] if len(t) > 1 else ""))
            elif mode == "TECHNICAL EXPERIENCE":
                for c in cells:
                    ps = [p for p in cell_paras(c) if ptxt(p).strip()]
                    if not ps:
                        continue
                    card = dict(title=ptxt(ps[0]).strip(), lines=[])
                    for p in ps[1:]:
                        ri = runs_info(p)
                        lab = ri[0][0].strip() if ri and ri[0][1] else None
                        txt = "".join(r[0] for r in ri[1:] if True).strip() if lab else ptxt(p).strip()
                        card["lines"].append((lab, txt))
                    m["tech_cards"].append(card)
            elif mode == "EDUCATION":
                for c in cells:
                    ps = [p for p in cell_paras(c) if ptxt(p).strip()]
                    if ps:
                        t0 = ptxt(ps[0])
                        title, date = (t0.split("\t", 1) + [""])[:2]
                        m["education"].append(dict(title=title.strip(), date=date.strip(),
                                                   school=ptxt(ps[1]).strip() if len(ps) > 1 else ""))
            elif mode == "TRAINING & PROFESSIONAL CERTIFICATIONS":
                for c in cells:
                    ps = [p for p in cell_paras(c) if ptxt(p).strip()]
                    if ps:
                        m["training"].append((ptxt(ps[0]).strip(), ptxt(ps[1]).strip() if len(ps) > 1 else ""))
    return m


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    import json
    m = parse()
    print(json.dumps(m, ensure_ascii=False, indent=1))
