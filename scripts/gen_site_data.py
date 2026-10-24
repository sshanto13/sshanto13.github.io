"""Generate src/data/cv.ts from the user's final CV docx (text copied verbatim)."""
import json, re, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")
from cv_model import parse

ROOT = Path(__file__).resolve().parent.parent
m = parse()


def clean(s):
    return s.replace("\u00a0", " ").strip() if s else s


COMPANY = {
    "DESTINIQUE TRAVEL LLC": ("Destinique Travel LLC", "https://destinique.com/", "img/logos/destinique.webp"),
    "JS GROUP": ("JS Group", "https://jsenterprise.com.bd/", "img/logos/js-enterprise.png"),
    "ADN TELECOM PLC": ("ADN Telecom PLC", "https://www.adntel.com.bd/", "img/logos/adn-telecom.png"),
    "NYBSYS INC.": ("NybSys Inc.", "https://nybsys.com/", "img/logos/nybsys.svg"),
    "DHAKA WEB HOST LTD.": ("Dhaka Web Host Ltd.", "https://www.dhakawebhost.com/", "img/logos/dhaka-web-host.png"),
    "BUZFI LLC": ("Buzfi LLC", None, "img/logos/buzfi.png"),
    "MOTMANS DIGITAL VISION": ("Motmans Digital Vision", "https://mdv-digital-services.be/", "img/logos/motmans.png"),
    "STCK.PRO": ("STCK.PRO", None, None),
    "SPIRETEC LIMITED": ("Spiretec Limited", "https://spiretecbd.com/", "img/logos/spiretec.png"),
}


def split_date(raw):
    parts = [p.strip() for p in raw.split("|")]
    engagement = None
    if parts and parts[0] == "Consultancy":
        engagement = "Consultancy"
        parts = parts[1:]
    period = location = None
    for p in parts:
        if re.search(r"\d{4}", p) or p.endswith("Present"):
            period = p
        else:
            location = p
    return engagement, period, location


def role(e, gov=False):
    eng, period, loc = split_date(clean(e["date"]))
    r = dict(title=clean(e["title"]), period=period, location=loc, engagement=eng)
    if gov:
        t, c = e["title"].split("\u2014", 1)
        r["title"] = clean(t)
        r["company"] = clean(c)
        r["companyUrl"] = "https://ictd.gov.bd/"
        r["logo"] = "img/logos/ictd.png"
    else:
        name, url, logo = COMPANY[clean(e["company"])]
        r["company"], r["companyUrl"], r["logo"] = name, url, logo
    r["intro"] = clean(e.get("intro"))
    bullets, kp = [], None
    for b in e["bullets"]:
        if b[0] == "label":
            kp = clean(b[2])
        else:
            bullets.append(clean(b[1]))
    r["bullets"] = bullets
    r["keyProjects"] = kp
    t = clean(e.get("tech"))
    r["stack"] = [x.strip() for x in t.rstrip(".").split(",")] if t else []
    return r


data = dict(
    name=m["name"],
    headline=m["headline"],
    summary=m["summary"],
    highlights=[dict(title=a, text=b) for a, b in m["highlights"]],
    techGroups=[dict(title=c["title"], lines=[dict(label=(l or "").rstrip(":"), text=t) for l, t in c["lines"]])
                for c in m["tech_cards"]],
    experience=[role(e) for e in m["experience"]],
    consultancy=[role(e) for e in m["consultancy"]],
    government=[role(e, gov=True) for e in m["gov"]],
    education=[dict(degree=e["title"], org=e["school"], time=e["date"]) for e in m["education"]],
    training=[],
    headings={k: v for k, v in m["headings"].items()},
)
for a, b in m["training"]:
    org, _, yr = b.rpartition("\u2014")
    data["training"].append(dict(degree=a, org=clean(org), time=clean(yr)))

# drop None values -> undefined in TS
def prune(o):
    if isinstance(o, dict):
        return {k: prune(v) for k, v in o.items() if v is not None}
    if isinstance(o, list):
        return [prune(x) for x in o]
    return o


js = json.dumps(prune(data), ensure_ascii=False, indent=2)
ts = (
    "// Generated from the final CV (Sumon_Ray_CV_Lead_IT_Digitalisation_2026.pdf.docx). Edit the CV, then re-run\n"
    "// scripts/gen_site_data.py - do not hand-edit text here.\n\n"
    "export type CvRole = {\n  title: string;\n  company: string;\n  companyUrl?: string;\n  logo?: string;\n"
    "  period?: string;\n  location?: string;\n  engagement?: string;\n  intro?: string;\n  bullets: string[];\n"
    "  keyProjects?: string;\n  stack: string[];\n};\n\n"
    f"export const cv = {js};\n"
)
(ROOT / "src" / "data" / "cv.ts").write_text(ts, encoding="utf-8")
print("written", len(ts))
