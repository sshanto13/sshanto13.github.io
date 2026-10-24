"""Typography-only pass on the user's edited Sumon_Ray_CV_Reviewed.docx.

Text is untouched except a handful of simple grammar fixes (listed in GRAMMAR).
Output: Sumon_Ray_CV_Reviewed_typography.docx / .pdf
"""
import copy, sys, os
from pathlib import Path
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, str(Path(__file__).parent))
from build_cv_reviewed import Measurer, fix_order   # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(r"c:\xampp\htdocs\sshanto13\sshanto13.github.io\public\files")
SRC = BASE / "Sumon_Ray_CV_Reviewed.docx"
OUT = BASE / "Sumon_Ray_CV_Reviewed_typography.docx"
SRC_COPY = Path(os.environ.get("TEMP", ".")) / "_reviewed_src_snapshot.docx"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SLATE, BODY = "475569", "334155"


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def runs(p):
    return [r for r in p.iter(qn("w:r")) if r.find(qn("w:t")) is not None]


def set_text(r, text):
    ts = r.findall(qn("w:t"))
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
        r.remove(t)


def set_rpr(r, font="Calibri", size=None, color=None, bold=None, italic=None):
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    for tag in ("w:rFonts", "w:b", "w:i", "w:color", "w:sz", "w:szCs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), font)
    rPr.append(rf)
    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if color:
        c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    if size:
        for tag in ("w:sz", "w:szCs"):
            s = OxmlElement(tag); s.set(qn("w:val"), str(int(size * 2))); rPr.append(s)


def apply_fixes(doc):
    body = doc.element.body
    paras = [e for e in body.iterchildren() if e.tag == W + "p"]
    log = []

    # ---- 0. stray empty paragraphs inside the cards (extra blank line at top/bottom of a card) ----
    tbls = [e for e in body.iterchildren() if e.tag == W + "tbl"]
    for tbl in tbls[1:]:
        for tc in tbl.iter(qn("w:tc")):
            ps = tc.findall(qn("w:p"))
            if any(ptext(p).strip() for p in ps):
                for p in ps:
                    if not ptext(p).strip():
                        tc.remove(p)
                        log.append("removed empty paragraph inside a card")

    # ---- 1. grammar (simple) -------------------------------------------------------------------
    for p in body.iter(qn("w:p")):
        t = ptext(p)
        rs = runs(p)
        if t.startswith("Florida, USA") and "boutique" in t:
            for r in rs:
                s = r.find(qn("w:t")).text
                if s.startswith("a boutique"):
                    set_text(r, "boutique,")
            log.append("Destinique intro: removed stray 'a' / double space")
        if t.startswith("Bangladeshi top web hosting"):
            set_text(rs[0], "Bangladesh-based")
            last = rs[-1]
            last_t = last.find(qn("w:t")).text
            set_text(last, last_t.replace("domains, and server solutions.", "domain, and server solutions provider."))
            log.append("Dhaka Web Host intro: 'Bangladesh-based ... solutions provider'")
        if t.startswith("Supported more than 10 business entity launch"):
            set_text(rs[0], "Supported the launch of ")
            for r in rs:
                s = r.find(qn("w:t")).text
                if s == " entity":
                    set_text(r, " entities")
                if s.startswith(" launch by coordinating"):
                    set_text(r, s.replace(" launch by", " by", 1))
            log.append("Spiretec bullet: 'Supported the launch of more than 10 business entities by ...'")
        for r in rs:
            tt = r.find(qn("w:t")).text or ""
            if "analyse" in tt:
                set_text(r, tt.replace("analyse", "analyze")); log.append("analyse -> analyze")
            if "order-fulfilment" in tt or "order fulfilment" in tt:
                set_text(r, tt.replace("order-fulfilment", "order-fulfillment").replace("order fulfilment", "order fulfillment"))
                log.append("fulfilment -> fulfillment (prose)")

    # ---- 2. punctuation/typography details ------------------------------------------------------
    # missing space before the '|' in date lines
    for p in body.iter(qn("w:p")):
        for r in runs(p):
            tt = r.find(qn("w:t")).text or ""
            if tt.startswith("|"):
                set_text(r, " " + tt)
                log.append("date line: space before '|'")
    # heading: trailing space run + mixed case ("Global PROJECT ENGAGEMENTS")
    for p in paras:
        t = ptext(p)
        if t.strip() == "TECHNICAL EXPERIENCE":
            rr = runs(p)
            if ptext(rr[-1].getparent()) != t.rstrip():
                pass
            if (rr[-1].find(qn("w:t")).text or "").strip() == "":
                rr[-1].getparent().remove(rr[-1])
            log.append("heading: trailing space removed")
        if t.startswith("Global PROJECT"):
            rr = runs(p)
            # same character formatting as the other section headings, rendered uppercase
            for r in rr:
                set_rpr(r, font="Arial", size=11, color="0F172A", bold=True)
                rPr = r.find(qn("w:rPr"))
                sp = rPr.find(qn("w:spacing"))
                if sp is None:
                    sp = OxmlElement("w:spacing"); rPr.append(sp)
                sp.set(qn("w:val"), "18")
                cp = OxmlElement("w:caps"); rPr.insert(1, cp)
            log.append("heading 'Global PROJECT ENGAGEMENTS' -> consistent caps heading style")

    # reference heading tracking from a normal heading
    ref_sp = None
    for p in paras:
        if ptext(p).strip() == "PROFESSIONAL EXPERIENCE":
            rr = runs(p)[0]
            sp = rr.find(qn("w:rPr")).find(qn("w:spacing"))
            ref_sp = sp.get(qn("w:val")) if sp is not None else None
    if ref_sp:
        for p in paras:
            if ptext(p).startswith("Global PROJECT") or ptext(p).startswith("TECHNICAL"):
                for r in runs(p):
                    rPr = r.find(qn("w:rPr"))
                    sp = rPr.find(qn("w:spacing"))
                    if sp is None:
                        sp = OxmlElement("w:spacing"); rPr.append(sp)
                    sp.set(qn("w:val"), ref_sp)

    # ---- 3. company intro lines: one style ------------------------------------------------------
    for i, p in enumerate(paras):
        prev = paras[i - 1] if i else None
        # intro line = italic 9pt paragraph straight after a company line
        rs = runs(p)
        if not rs or prev is None:
            continue
        r0pr = rs[0].find(qn("w:rPr"))
        italic = r0pr is not None and r0pr.find(qn("w:i")) is not None
        sz = r0pr.find(qn("w:sz")) if r0pr is not None else None
        pstyle = p.find(qn("w:pPr")).find(qn("w:pStyle")) if p.find(qn("w:pPr")) is not None else None
        if italic and pstyle is None and sz is not None and sz.get(qn("w:val")) == "18" and ptext(prev).strip() and \
           (ptext(prev).strip().isupper()):
            for r in rs:
                set_rpr(r, font="Calibri", size=9, color=SLATE, italic=True)
            last = rs[-1]
            lt = last.find(qn("w:t")).text
            if not lt.rstrip().endswith("."):
                set_text(last, lt.rstrip() + ".")
                log.append("intro: terminal period added")
    # remove stray empty paragraph(s) between intro line and first bullet
    for i in range(len(paras) - 1, 0, -1):
        p = paras[i]
        if not ptext(p).strip() and p.find(qn("w:pPr")) is not None:
            sp = p.find(qn("w:pPr")).find(qn("w:spacing"))
            line = sp.get(qn("w:line")) if sp is not None else None
            nxt = paras[i + 1] if i + 1 < len(paras) else None
            prv = paras[i - 1]
            # an empty 12pt-ish paragraph directly after an italic intro (not a spacer line)
            if line in ("241",) and nxt is not None and ptext(nxt).strip() and prv is not None and "ListBullet" not in str(nxt):
                body.remove(p); log.append("removed stray empty paragraph")

    # ---- 4. consultancy bullets -> same bullet formatting as the rest -------------------------------
    ref_bullet = None
    for p in paras:
        if ptext(p).startswith("Develop and maintain web applications"):
            ref_bullet = p
    refpPr = ref_bullet.find(qn("w:pPr"))
    bl = []
    for p in list(body.iterchildren()):
        if p.tag != W + "p":
            continue
        pPr = p.find(qn("w:pPr"))
        st = pPr.find(qn("w:pStyle")) if pPr is not None else None
        if st is not None and st.get(qn("w:val")) == "NormalWeb":
            new = copy.deepcopy(refpPr)
            p.replace(pPr, new)
            for r in runs(p):
                set_rpr(r, font="Calibri", size=9.5, color=BODY)
            bl.append(p)
            log.append("bullet restyled: " + ptext(p)[:40])
    # keep-with-next chain for the consultancy blocks
    allp = [e for e in body.iterchildren() if e.tag == W + "p"]
    for i, p in enumerate(allp):
        if p in bl:
            pPr = p.find(qn("w:pPr"))
            for k in pPr.findall(qn("w:keepNext")):
                pPr.remove(k)
            prev_is_bullet = allp[i - 1] in bl
            nxt = allp[i + 1] if i + 1 < len(allp) else None
            next_is_tech = nxt is not None and ptext(nxt).strip().startswith("Technologies")
            if (not prev_is_bullet) or next_is_tech:
                kn = OxmlElement("w:keepNext"); pPr.insert(1, kn)
    fix_order(body)
    return log


def spacing_scale(doc, t):
    """Scale vertical gaps (before/after + spacer lines) of top-level paragraphs by t."""
    for p in doc.element.body.iterchildren():
        if p.tag != W + "p":
            continue
        pPr = p.find(qn("w:pPr"))
        sp = pPr.find(qn("w:spacing")) if pPr is not None else None
        if sp is None:
            continue
        for a in ("before", "after"):
            v = sp.get(qn("w:" + a))
            if v is not None and (sp.get(qn("w:" + a + "_orig")) is None):
                sp.set(qn("w:" + a), str(int(round(int(v) * t))))
        if not ptext(p).strip() and sp.get(qn("w:lineRule")) == "exact":
            sp.set(qn("w:line"), str(max(20, int(round(int(sp.get(qn("w:line"))) * t)))))


def load():
    return docx.Document(str(SRC_COPY))


def scale_range(doc, lo, hi, t):
    kids = list(doc.element.body.iterchildren())
    for p in kids[lo:hi]:
        if p.tag != W + "p":
            continue
        pPr = p.find(qn("w:pPr"))
        sp = pPr.find(qn("w:spacing")) if pPr is not None else None
        if sp is None:
            continue
        for a in ("before", "after"):
            v = sp.get(qn("w:" + a))
            if v is not None:
                sp.set(qn("w:" + a), str(int(round(int(v) * t))))
        if not ptext(p).strip() and sp.get(qn("w:lineRule")) == "exact":
            sp.set(qn("w:line"), str(max(20, int(round(int(sp.get(qn("w:line"))) * t)))))


def card4_end(doc):
    for i, e in enumerate(doc.element.body.iterchildren()):
        if e.tag == W + "tbl" and "IT Infrastructure, Cloud" in ptext(e):
            return i + 1
    raise RuntimeError("card 4 not found")


def build_variant(t1, t2, nbsp):
    from docx.text.paragraph import Paragraph
    from build_cv_reviewed import join_last_two_words
    d = load()
    apply_fixes(d)
    c4 = card4_end(d)
    scale_range(d, 0, c4, t1)
    scale_range(d, c4, 10 ** 6, t2)
    if nbsp:
        for pe in list(d.element.body.iter(qn("w:p"))):
            para = Paragraph(pe, d._body)
            txt = para.text.rstrip()
            for (a, b) in nbsp:
                if txt.endswith(a + " " + b):
                    join_last_two_words(para)
    fix_order(d.element)
    return d


if __name__ == "__main__":
    import shutil
    if not SRC_COPY.exists() or "--refresh" in sys.argv:
        shutil.copy(SRC, SRC_COPY)
    mode = sys.argv[1] if len(sys.argv) > 1 else "probe"
    m = Measurer()
    try:
        if mode == "probe":
            d = load(); log = apply_fixes(d)
            print("\n".join(sorted(set(log))))
            for t in (0.5, 0.7, 0.85, 1.0, 1.2):
                d = load(); apply_fixes(d); spacing_scale(d, t)
                pg, info = m.measure(d)
                print("t", t, "pages", pg, [(round(i["bottom"]), i["first"][:1]) for i in info])
        elif mode == "final":
            nbsp = []

            def p1_ok(t1):
                pg, inf = m.measure(build_variant(t1, 1.0, nbsp))
                return len(inf) > 1 and inf[1]["first"][0].startswith("5. Leadership")

            lo, hi = 0.35, 1.0
            if p1_ok(1.0):
                lo = 1.0
            else:
                for _ in range(7):
                    mid = (lo + hi) / 2
                    if p1_ok(mid):
                        lo = mid
                    else:
                        hi = mid
            t1 = lo
            print("t1", t1, "p1_ok", p1_ok(t1))

            def p2_ok(t2):
                pg, inf = m.measure(build_variant(t1, t2, nbsp))
                return pg == 4 and inf[-1]["bottom"] <= 772, inf

            lo, hi = 0.6, 1.8
            for _ in range(8):
                mid = (lo + hi) / 2
                ok, _ = p2_ok(mid)
                if ok:
                    lo = mid
                else:
                    hi = mid
            t2 = lo
            for _ in range(3):
                pg, inf = m.measure(build_variant(t1, t2, nbsp))
                wd = [x for pgi in inf for x in pgi["widows"] if x not in nbsp]
                print("t2", round(t2, 3), "pages", pg, [(round(i["bottom"]), i["first"][:1]) for i in inf], "widows", wd)
                if not wd:
                    break
                nbsp += wd
            d = build_variant(t1, t2, nbsp)
            out = OUT
            try:
                d.save(out)
            except PermissionError:
                out = OUT.with_name(OUT.stem + "_v2.docx"); d.save(out)
            dd = m.app.Documents.Open(str(out), ReadOnly=True)
            dd.SaveAs2(str(out.with_suffix(".pdf")), 17)
            print("pages", dd.ComputeStatistics(2))
            dd.Close(False)
            print("saved", out)
    finally:
        m.app.Quit()
