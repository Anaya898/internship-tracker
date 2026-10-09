#!/usr/bin/env python3
"""Daily internship tracker.

Scrapes the GitHub internship lists, merges them into an Excel workbook (one sheet
per role, Applied/Pending column that is NEVER overwritten) and rebuilds a static site.

  python update_tracker.py --xlsx "C:/path/Internship_Tracker.xlsx" --site docs
"""
import argparse, datetime as dt, hashlib, json, os, re, shutil
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SOURCES = [
    dict(key="simplify", name="SimplifyJobs / Pitt CSC", repo="SimplifyJobs/Summer2027-Internships", branch="dev", kind="html"),
    dict(key="vansh", name="Vansh & Ouckah (CSCareers)", repo="vanshb03/Summer2027-Internships", branch="dev", kind="md", ncols=5),
    dict(key="speedy", name="SpeedyApply AI", repo="speedyapply/2027-AI-College-Jobs", branch="main", kind="md", ncols=6),
]
EXTRA_LINKS = [
    ("Simplify - Inactive listings", "https://github.com/SimplifyJobs/Summer2027-Internships/blob/dev/README-Inactive.md"),
    ("Simplify - Off-season 2027", "https://github.com/SimplifyJobs/Summer2027-Internships/blob/dev/README-Off-Season.md"),
    ("Simplify - New grad repo", "https://github.com/SimplifyJobs/New-Grad-Positions"),
    ("Simplify - browser autofill", "https://simplify.jobs/install"),
    ("Vansh & Ouckah - Off-season", "https://github.com/vanshb03/Summer2027-Internships/blob/dev/OFFSEASON_README.md"),
    ("Vansh & Ouckah - New grad", "https://github.com/cvrve/New-Grad"),
    ("CSCareers Discord", "https://discord.gg/cscareers"),
    ("SWEList email alerts", "https://swelist.com"),
]
ROLES = ["Business Analyst", "Data Sci-AI-ML", "Product Mgmt", "Software Eng", "Quant Finance", "Hardware Eng", "Other"]
COLS = ["Status", "Company", "Role", "Location", "Date Posted", "Apply", "Simplify", "Company Site", "Salary",
        "Flags", "Source", "First Seen", "Last Seen", "Still Listed", "Notes", "ID"]
LOG_COLS = ["Run (UTC)", "Source", "Repo", "Branch", "Last commit (UTC)", "Commit", "Changed since last run",
            "Rows scraped", "New rows", "Commit link"]
FLAGS = {"🛂": "No sponsorship", "🇺🇸": "US citizens only", "🔒": "Closed", "🔥": "FAANG+", "🎓": "Advanced degree"}
H = {"User-Agent": "internship-tracker"}
TODAY = dt.date.today()
NOW = dt.datetime.now(dt.timezone.utc)


# ---------- helpers ----------
def get(url, **kw):
    r = requests.get(url, headers=H, timeout=90, **kw)
    r.raise_for_status()
    return r


def clean_url(u):
    if not u or not u.startswith("http"):
        return ""
    p = urlsplit(u.strip())
    q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True) if not (k.startswith("utm_") or k in ("ref", "source"))]
    return urlunsplit((p.scheme, p.netloc, p.path.rstrip("/") or "/", urlencode(q), ""))


def soup(html):
    return BeautifulSoup(re.sub(r"</?br\s*/?>", "\n", html), "html.parser")


def txt(node, sep=" "):
    for s in node.find_all("summary"):
        s.decompose()
    parts = [x.strip() for x in node.get_text("\n").split("\n")]
    return sep.join(p for p in parts if p)


def age_to_date(s):
    m = re.match(r"(\d+)\s*(mo|d|w|y)", (s or "").strip().lower())
    if not m:
        return ""
    n, u = int(m.group(1)), m.group(2)
    return (TODAY - dt.timedelta(days={"d": 1, "w": 7, "mo": 30, "y": 365}[u] * n)).isoformat()


def md_to_date(s):
    try:
        d = dt.datetime.strptime(f"{s.strip()} {TODAY.year}", "%b %d %Y").date()
        return (d if d <= TODAY else d.replace(year=d.year - 1)).isoformat()
    except ValueError:
        return ""


RULES = [
    ("Quant Finance", r"quant|trader|trading"),
    ("Product Mgmt", r"product (management|manager|intern|owner|associate)|\bapm\b|program manag"),
    ("Business Analyst", r"\banalyst\b|business intelligence|\bbi\b|strategy|consult"),
    ("Data Sci-AI-ML", r"data scien|machine learning|\bml\b|\bai\b|artificial|deep learning|nlp|applied scien|research scien|data engineer|\bllm"),
    ("Hardware Eng", r"hardware|electrical|firmware|embedded|fpga|asic|mechanical|circuit|silicon|vlsi|\brf\b"),
    ("Software Eng", r"software|developer|engineer|swe|full.?stack|front.?end|back.?end|\bweb\b|mobile|ios|android|devops|cloud|security|\bsre\b"),
]


def classify(role, hint=None, default="Other"):
    t = role.lower()
    if hint in ("Hardware Eng", "Quant Finance", "Product Mgmt"):
        return hint
    if re.search(RULES[2][1], t) and not re.search(r"software|engineer|developer|scientist", t):
        return "Business Analyst"
    if hint:
        return hint
    for name, pat in RULES:
        if re.search(pat, t):
            return name
    return default


def mk(src, company, role, loc, salary, posted, apply_, simp, site, hint=None, default="Other"):
    flags = []
    for k, v in FLAGS.items():
        if k in role or k in company:
            flags.append(v)
        role, company = role.replace(k, ""), company.replace(k, "")
    fix = lambda s: re.sub(r"\s+", " ", s.replace("\ufe0f", "")).strip()
    role, company = fix(role), fix(company)
    return dict(company=company, role=role, location=loc, salary=salary, posted=posted, apply=apply_, simplify=simp,
                site=site, flags=", ".join(flags), sources={src}, cat=classify(role, hint, default))


def split_links(td):
    apply_, simp = "", ""
    for a in td.find_all("a", href=True):
        u = clean_url(a["href"])
        if "simplify.jobs/p/" in u:
            simp = simp or u
        elif u and not apply_:
            apply_ = u
    return apply_, simp


# ---------- parsers ----------
def parse_simplify(text, src):
    sec = {"Software": "Software Eng", "Product": "Product Mgmt", "Data Science": "Data Sci-AI-ML",
           "Quantitative": "Quant Finance", "Hardware": "Hardware Eng"}
    out = []
    for chunk in re.split(r"(?m)^## ", text)[1:]:
        hint = next((v for k, v in sec.items() if k in chunk.split("\n", 1)[0]), None)
        if not hint:
            continue
        prev = ("", "")
        for tr in soup(chunk).find_all("tr"):
            tds = tr.find_all("td", recursive=False)
            if len(tds) < 5:
                continue
            name = txt(tds[0])
            if name == "↳":
                company, site = prev
            else:
                a = tds[0].find("a", href=True)
                company, site = name, clean_url(a["href"]) if a else ""
                prev = (company, site)
            apply_, simp = split_links(tds[3])
            out.append(mk(src, company, txt(tds[1]), txt(tds[2], "; "), "", age_to_date(txt(tds[4])), apply_, simp, site, hint))
    return out


def parse_md(text, src, ncols, key):
    if key == "vansh":
        text = text.split("## We love our contributors")[0]
    out, prev, hint, n = [], ("", ""), None, ncols
    for line in text.split("\n"):
        if line.startswith("### "):
            hint = "Quant Finance" if "Quant" in line else None
        if not line.startswith("|") or re.match(r"^\|\s*-", line) or line.startswith("| Company"):
            continue
        c = [x.strip() for x in line.strip().strip("|").split("|")]
        if len(c) < n:
            continue
        if len(c) > n:
            c = c[:1] + [" | ".join(c[1:len(c) - n + 2])] + c[len(c) - n + 2:]
        cs = soup(c[0])
        name = txt(cs)
        if name == "↳":
            company, site = prev
        else:
            a = cs.find("a", href=True)
            company, site = name, clean_url(a["href"]) if a else ""
            prev = (company, site)
        apply_, simp = split_links(soup(c[3] if key == "vansh" else c[4]))
        if key == "vansh":
            rec = mk(src, company, txt(soup(c[1])), txt(soup(c[2]), "; "), "", md_to_date(c[4]), apply_, simp, site, hint, "Software Eng")
        else:
            rec = mk(src, company, txt(soup(c[1])), txt(soup(c[2]), "; "), txt(soup(c[3])), age_to_date(c[5]), apply_, simp, site, hint, "Data Sci-AI-ML")
        if rec["company"] and rec["role"]:
            out.append(rec)
    return out


def rec_id(r):
    base = r["apply"] or f'{r["company"]}|{r["role"]}|{r["location"]}'
    return hashlib.sha1(base.lower().encode()).hexdigest()[:12]


# ---------- GitHub update info ----------
def commit_info(src):
    h = dict(H)
    if os.environ.get("GITHUB_TOKEN"):
        h["Authorization"] = "Bearer " + os.environ["GITHUB_TOKEN"]
    try:
        r = requests.get(f"https://api.github.com/repos/{src['repo']}/commits", headers=h, timeout=30,
                         params={"sha": src["branch"], "path": "README.md", "per_page": 1})
        r.raise_for_status()
        j = r.json()[0]
        return dict(sha=j["sha"][:7], date=j["commit"]["committer"]["date"].replace("T", " ").rstrip("Z"), url=j["html_url"])
    except Exception as e:  # noqa  (API rate-limited -> fall back to the public commit feed)
        try:
            x = get(f"https://github.com/{src['repo']}/commits/{src['branch']}/README.md.atom").text
            e1 = x.split("<entry>")[1]
            url = re.search(r'href="([^"]+/commit/[0-9a-f]+)"', e1).group(1)
            when = re.search(r"<updated>([^<]+)</updated>", e1).group(1)
            return dict(sha=url.rsplit("/", 1)[1][:7], date=when.replace("T", " ").rstrip("Z"), url=url)
        except Exception as e2:  # noqa
            print("commit lookup failed:", src["repo"], e, e2)
            return dict(sha="", date="", url="")


# ---------- workbook state ----------
def load_state(path):
    state, log, wb = {}, [], None
    if path.exists():
        wb = load_workbook(path)
        for ws in wb.worksheets:
            if ws.title in ROLES:
                head = [c.value for c in ws[1]]
                if "ID" not in head:
                    continue
                ix = {h: i for i, h in enumerate(head)}
                for row in ws.iter_rows(min_row=2):
                    g = lambda k: row[ix[k]].value if k in ix and ix[k] < len(row) else None
                    lk = lambda k: (row[ix[k]].hyperlink.target if row[ix[k]].hyperlink else (g(k) if str(g(k) or "").startswith("http") else "")) if k in ix else ""
                    if not g("ID"):
                        continue
                    state[g("ID")] = dict(
                        id=g("ID"), cat=ws.title, status=g("Status") or "Pending", notes=g("Notes") or "",
                        company=g("Company") or "", role=g("Role") or "", location=g("Location") or "",
                        posted=str(g("Date Posted") or ""), apply=lk("Apply"), simplify=lk("Simplify"), site=lk("Company Site"),
                        salary=g("Salary") or "", flags=g("Flags") or "", sources=set(str(g("Source") or "").split(", ")),
                        first_seen=str(g("First Seen") or ""), last_seen=str(g("Last Seen") or ""))
            elif ws.title == "Update Log":
                log = [[c.value for c in r] for r in ws.iter_rows(min_row=2) if r[0].value]
    return state, log, wb


# ---------- Excel ----------
def write_sheet(wb, title, header, rows, widths, link_cols=(), hide=(), validate=False):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title, len(wb.sheetnames) if title in ("Update Log", "Sources") else ROLES.index(title) if title in ROLES else None)
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(vertical="center")
    for r in rows:
        ws.append([v if not isinstance(v, tuple) else v[1] for v in r])
        for i, v in enumerate(r):
            if isinstance(v, tuple) and v[0]:
                cell = ws.cell(row=ws.max_row, column=i + 1)
                cell.hyperlink, cell.font = v[0], Font(color="0563C1", underline="single")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for h in hide:
        ws.column_dimensions[h].hidden = True
    ws.freeze_panes = "C2" if validate else "A2"
    ws.auto_filter.ref = ws.dimensions
    if validate and ws.max_row > 1:
        dv = DataValidation(type="list", formula1='"Pending,Applied"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(f"A2:A{ws.max_row}")
        rng = f"A2:A{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Applied"'], fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Pending"'], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default="tracker/Internship_Tracker.xlsx")
    ap.add_argument("--site", default="docs")
    a = ap.parse_args()
    path = Path(a.xlsx)
    path.parent.mkdir(parents=True, exist_ok=True)

    fresh, counts, infos = {}, {}, {}
    for s in SOURCES:
        try:
            text = get(f"https://raw.githubusercontent.com/{s['repo']}/{s['branch']}/README.md").text
            recs = parse_simplify(text, s["name"]) if s["kind"] == "html" else parse_md(text, s["name"], s["ncols"], s["key"])
        except Exception as e:  # noqa
            print("FAILED", s["repo"], e)
            recs = []
        counts[s["key"]], infos[s["key"]] = len(recs), commit_info(s)
        for r in recs:
            i = rec_id(r)
            if i in fresh:
                o = fresh[i]
                o["sources"] |= r["sources"]
                for k in ("salary", "simplify", "site", "posted"):
                    o[k] = o[k] or r[k]
            else:
                fresh[i] = r
    if not fresh:
        raise SystemExit("Nothing scraped - leaving workbook untouched.")

    state, log, wb = load_state(path)
    wb = wb or Workbook()
    if wb.sheetnames == ["Sheet"]:
        del wb["Sheet"]

    final, new_by_src = {}, {k: 0 for k in counts}
    for i, r in fresh.items():
        old = state.get(i)
        r.update(id=i, status=old["status"] if old else "Pending", notes=old["notes"] if old else "",
                 first_seen=(old["first_seen"] if old and old["first_seen"] else TODAY.isoformat()),
                 last_seen=TODAY.isoformat(), listed="Yes", sources=sorted(r["sources"]))
        if not old:
            for s in SOURCES:
                if s["name"] in r["sources"]:
                    new_by_src[s["key"]] += 1
        final[i] = r
    for i, o in state.items():  # rows that disappeared from the lists
        if i in fresh:
            continue
        age = (TODAY - dt.date.fromisoformat(o["last_seen"])).days if re.match(r"\d{4}-", o["last_seen"]) else 0
        if o["status"] == "Applied" or age <= 30:
            o.update(listed="No", sources=sorted(x for x in o["sources"] if x))
            final[i] = o

    by_cat = {r: [] for r in ROLES}
    for r in final.values():
        by_cat.setdefault(r["cat"] if r["cat"] in by_cat else "Other", []).append(r)
    for name in ROLES:
        rows = sorted(by_cat[name], key=lambda r: (r["posted"] or "", r["company"]), reverse=True)
        write_sheet(wb, name, COLS, [[
            r["status"], r["company"], r["role"], r["location"], r["posted"],
            (r["apply"], "Apply"), (r["simplify"], "Simplify" if r["simplify"] else ""), (r["site"], "Site" if r["site"] else ""),
            r["salary"], r["flags"], ", ".join(r["sources"]), r["first_seen"], r["last_seen"], r["listed"], r["notes"], r["id"]]
            for r in rows],
            [10, 28, 46, 26, 12, 9, 10, 8, 10, 22, 26, 12, 12, 10, 30, 14], hide=("P",), validate=True)
        by_cat[name] = rows

    # update log (newest first)
    last = {}
    for row in reversed(log):
        last.setdefault(row[1], row[5])
    run = NOW.strftime("%Y-%m-%d %H:%M")
    for s in SOURCES:
        c = infos[s["key"]]
        if not c["sha"] or last.get(s["name"]) == c["sha"]:
            continue  # nothing new on GitHub since the last logged commit -> no repeated log rows
        log.append([run, s["name"], s["repo"], s["branch"], c["date"], c["sha"],
                    ("Yes" if last.get(s["name"]) != c["sha"] else "No") if c["sha"] else "Unknown", counts[s["key"]], new_by_src[s["key"]], c["url"]])
    log = sorted(log, key=lambda r: str(r[0]), reverse=True)[:600]
    write_sheet(wb, "Update Log", LOG_COLS, [r[:9] + [(r[9], "commit") if r[9] else ""] for r in log], [18, 28, 38, 8, 20, 10, 12, 12, 10, 10])
    write_sheet(wb, "Sources", ["Name", "Link"],
                [[s["name"], (f"https://github.com/{s['repo']}/blob/{s['branch']}/README.md", f"github.com/{s['repo']}")] for s in SOURCES]
                + [[n, (u, u)] for n, u in EXTRA_LINKS], [40, 90])
    wb.save(path)

    # website
    site = Path(a.site)
    site.mkdir(parents=True, exist_ok=True)
    keys = dict(id="id", st="status", co="company", ro="role", lo="location", dt="posted", ap="apply", si="simplify", ws="site",
                sa="salary", fl="flags", fs="first_seen", ls="listed")
    data = dict(
        generated=NOW.strftime("%Y-%m-%d %H:%M UTC"), xlsx=path.name,
        roles=[dict(name=n, rows=[{k: r[v] for k, v in keys.items()} | {"so": ", ".join(r["sources"])} for r in by_cat[n]]) for n in ROLES],
        sources=[dict(name=s["name"], url=f"https://github.com/{s['repo']}", commit=infos[s["key"]], rows=counts[s["key"]],
                      new=new_by_src[s["key"]]) for s in SOURCES],
        links=[dict(n=n, u=u) for n, u in EXTRA_LINKS])
    html = (Path(__file__).parent / "site_template.html").read_text(encoding="utf-8")
    (site / "index.html").write_text(html.replace("__DATA__", json.dumps(data, ensure_ascii=False).replace("</", "<\\/")), encoding="utf-8")
    shutil.copy(path, site / path.name)
    print(f"OK: {len(final)} rows | " + " | ".join(f"{n}: {len(by_cat[n])}" for n in ROLES))


if __name__ == "__main__":
    main()
