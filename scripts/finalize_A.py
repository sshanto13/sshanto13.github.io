import sys
from pathlib import Path
import fitz
sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")
from cv_model import parse
import cv_layouts as L
from build_cv_reviewed import Measurer, fix_order

BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
OUT = BASE / "Sumon_Ray_CV_ACI_Lead_IT_Final.docx"
m = parse()
app = Measurer()


def widows_of(pdf_path):
    pd = fitz.open(pdf_path)
    res = []
    for pg in pd:
        lines = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if t:
                    lines.append((l["bbox"][1], l["bbox"][3], l["bbox"][0], l["bbox"][2], t))
        lines.sort()
        for i in range(1, len(lines)):
            y0, y1, x0, x1, t = lines[i]
            py0, py1, px0, px1, pt = lines[i - 1]
            if (y0 < 740 and len(t.split()) == 1 and len(t) > 3 and 0 < y0 - py0 < 17 and px1 > 470
                    and abs(px0 - x0) < 20 and x0 < 95 and not t.endswith(":")):
                res.append((pt.split()[-1], t))
    return res


def build(rg, hg, nbsp):
    doc = L.layout_A(m, RG=rg, HG=hg, nbsp=nbsp)
    fix_order(doc.element)
    fix_order(doc.sections[0].footer._element)
    return doc


def measure(doc, tag="t"):
    p = Path(__import__("tempfile").gettempdir()) / "_fa.docx"
    doc.save(p)
    d = app.app.Documents.Open(str(p), ReadOnly=True)
    pdf = p.with_suffix(".pdf")
    d.SaveAs2(str(pdf), 17)
    n = d.ComputeStatistics(2)
    d.Close(False)
    pd = fitz.open(pdf)
    info = []
    for pg in pd:
        ls = [(l["bbox"][1], l["bbox"][3], "".join(s["text"] for s in l["spans"]).strip())
              for b in pg.get_text("dict")["blocks"] for l in b.get("lines", [])]
        ls = [x for x in ls if x[2] and x[0] < 742]
        ls.sort()
        info.append((round(max(x[1] for x in ls)), ls[0][2][:40]))
    w = widows_of(pdf)
    pd.close()
    return n, info, w


try:
    nbsp = []
    chosen = None
    for s in (0.95, 0.92, 0.9, 0.87, 0.85):
        n, info, w = measure(build(s, s, nbsp))
        print("scale", s, "pages", n, info)
        if n == 4 and info[2][1].startswith("Assistant Manager"):
            chosen = s
            break
    if chosen is None:
        raise SystemExit("no good scale")
    for _ in range(4):
        n, info, w = measure(build(chosen, chosen, nbsp))
        new = [x for x in w if x not in nbsp]
        print("widow round", n, info, "new", new)
        if not new:
            break
        nbsp += new
    doc = build(chosen, chosen, nbsp)
    doc.save(OUT)
    d = app.app.Documents.Open(str(OUT), ReadOnly=True)
    d.SaveAs2(str(OUT.with_suffix(".pdf")), 17)
    print("final pages", d.ComputeStatistics(2))
    d.Close(False)
    print("saved", OUT, "nbsp", nbsp)
finally:
    app.app.Quit()
