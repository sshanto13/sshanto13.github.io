import sys, docx
from docx.oxml.ns import qn
sys.stdout.reconfigure(encoding="utf-8")
path = sys.argv[1]
d = docx.Document(path)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def pinfo(p, tag=""):
    pPr = p.find(qn("w:pPr"))
    sp = ind = kn = None
    sty = None
    if pPr is not None:
        s = pPr.find(qn("w:spacing")); sp = dict(s.attrib) if s is not None else None
        if sp: sp = {k.split('}')[1]: v for k, v in sp.items()}
        i = pPr.find(qn("w:ind")); ind = {k.split('}')[1]: v for k, v in i.attrib.items()} if i is not None else None
        st = pPr.find(qn("w:pStyle")); sty = st.get(qn("w:val")) if st is not None else None
        kn = [c.tag.split('}')[1] for c in pPr if c.tag.split('}')[1] in ("keepNext", "keepLines", "pageBreakBefore", "numPr")]
    runs = []
    for r in p.iter(qn("w:r")):
        t = "".join(x.text or "" for x in r.iter(qn("w:t")))
        rPr = r.find(qn("w:rPr"))
        f = sz = col = b = it = None
        if rPr is not None:
            rf = rPr.find(qn("w:rFonts")); f = rf.get(qn("w:ascii")) if rf is not None else None
            s = rPr.find(qn("w:sz")); sz = int(s.get(qn("w:val"))) / 2 if s is not None else None
            c = rPr.find(qn("w:color")); col = c.get(qn("w:val")) if c is not None else None
            b = rPr.find(qn("w:b")) is not None; it = rPr.find(qn("w:i")) is not None
        runs.append((t[:28], f, sz, col, "B" if b else "", "I" if it else ""))
    txt = "".join(x.text or "" for x in p.iter(qn("w:t")))
    print(tag, "sty=", sty, "sp=", sp, "ind=", ind, kn)
    for r in runs:
        print("      ", r)


n = 0
for el in d.element.body.iterchildren():
    if el.tag == W + "p":
        pinfo(el, f"[{n}] P")
    elif el.tag == W + "tbl":
        print(f"[{n}] TABLE")
        for tc in el.iter(qn("w:tc")):
            for p in tc.findall(qn("w:p")):
                pinfo(p, "   cellP")
    n += 1
