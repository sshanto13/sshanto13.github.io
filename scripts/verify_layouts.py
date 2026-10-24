import sys, re, glob
from collections import Counter
from pathlib import Path
import docx
from docx.oxml.ns import qn
sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")


def tokens(path):
    d = docx.Document(str(path))
    txt = []
    for p in d.element.body.iter(qn("w:p")):
        t = "".join((x.text or "") if x.tag == qn("w:t") else " " for x in p.iter() if x.tag in (qn("w:t"), qn("w:tab")) and x.getparent().tag != qn("w:tabs"))
        txt.append(t)
    s = " ".join(txt).replace("\xa0", " ")
    s = re.sub(r"[\U00010000-\U0010FFFF\u2600-\u27BF\uFE0F\u20E3]", " ", s)   # emoji icons
    return Counter(re.findall(r"\S+", s))


src = tokens(BASE / "Sumon_Ray_CV_Reviewed_typography.docx")
print("source tokens", sum(src.values()))
ok = True
for f in sorted(glob.glob(str(BASE / "cv_layouts" / "*.docx"))):
    t = tokens(f)
    missing = src - t
    extra = t - src
    print(Path(f).name, "tokens", sum(t.values()), "missing", dict(missing), "extra", dict(extra))
