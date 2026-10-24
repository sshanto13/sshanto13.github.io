import sys, os
from pathlib import Path
import fitz
sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")
from cv_model import parse
import cv_layouts as L
from build_cv_reviewed import Measurer, fix_order

OUTDIR = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files\cv_layouts")
OUTDIR.mkdir(exist_ok=True)
PREV = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\_prev")
PREV.mkdir(exist_ok=True)

names = sys.argv[1:] or list(L.LAYOUTS)
m = parse()
app = Measurer()
try:
    for key in names:
        doc = L.LAYOUTS[key](m)
        fix_order(doc.element)
        try:
            fix_order(doc.sections[0].footer._element)
        except Exception:
            pass
        out = OUTDIR / f"Sumon_Ray_CV_{key}.docx"
        doc.save(out)
        d = app.app.Documents.Open(str(out), ReadOnly=True)
        pdf = out.with_suffix(".pdf")
        d.SaveAs2(str(pdf), 17)
        n = d.ComputeStatistics(2)
        d.Close(False)
        pd = fitz.open(pdf)
        bottoms = []
        for i, pg in enumerate(pd):
            ys = [l["bbox"][3] for b in pg.get_text("dict")["blocks"] for l in b.get("lines", []) if l["bbox"][1] < 792 - 50 and "".join(s["text"] for s in l["spans"]).strip()]
            bottoms.append(round(max(ys)) if ys else 0)
            pg.get_pixmap(dpi=80).save(PREV / f"{key}_{i + 1}.png")
        print(key, "pages", n, "bottoms", bottoms)
finally:
    app.app.Quit()
