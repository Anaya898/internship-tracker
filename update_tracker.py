#!/usr/bin/env python3
"""Internship tracker -> Google Sheet.

Scrapes the GitHub internship lists and syncs them into a Google Sheet (one tab per role).
Your Status / Notes columns are read back from the Sheet on every run and never overwritten.

  python update_tracker.py --sheet-id <ID> [--import-xlsx old.xlsx]   (needs env GOOGLE_SA_JSON)
  python update_tracker.py --dry-run                                  (scrape only, touches nothing)
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
COLS = ["Status", "New", "Company", "Role", "Location", "Date Posted", "Apply", "Simplify", "Company Site", "Salary",
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


# ---------- Google Sheets ----------
A1 = lambda t, rng: "'" + t.replace("'", "''") + "'!" + rng
TEXT_COLS = [2, 3, 4, 5, 9, 10, 11, 12, 13, 15, 16]          # kept as plain text (dates, ids, notes...)
WIDTHS = [80, 55, 200, 320, 170, 90, 70, 80, 100, 80, 130, 170, 90, 90, 80, 220, 110]
LOG_TEXT, LOG_W = [0, 1, 2, 3, 4, 5, 6], [130, 200, 260, 60, 140, 70, 110, 90, 70, 80]
STATUSES = ["Pending", "Applied", "Skip"]


def rgb(h):
    return {k: int(h[i:i + 2], 16) / 255 for k, i in (("red", 0), ("green", 2), ("blue", 4))}


def meta(sh):
    return {x["properties"]["title"]: x["properties"] for x in sh.fetch_sheet_metadata()["sheets"]}


def ensure_tab(sh, title, ncols, widths, text_cols, role=False):
    """Return (sheetId, rowCount); creates and formats the tab the first time."""
    props = meta(sh).get(title)
    if props:
        return props["sheetId"], props["gridProperties"]["rowCount"]
    sh.batch_update({"requests": [{"addSheet": {"properties": {"title": title, "gridProperties": {
        "rowCount": 6000, "columnCount": max(ncols, 10), "frozenRowCount": 1}}}}]})
    sid = meta(sh)[title]["sheetId"]
    rng = lambda i, r0=1: {"sheetId": sid, "startRowIndex": r0, "endRowIndex": 6000, "startColumnIndex": i, "endColumnIndex": i + 1}
    dim = lambda i, props, f: {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1}, "properties": props, "fields": f}}
    reqs = [
        {"repeatCell": {"range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {"userEnteredFormat": {"backgroundColor": rgb("1F3864"), "textFormat": {"bold": True, "foregroundColor": rgb("FFFFFF")}}},
                        "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        {"setBasicFilter": {"filter": {"range": {"sheetId": sid, "startRowIndex": 0, "startColumnIndex": 0, "endColumnIndex": ncols}}}},
    ]
    reqs += [{"repeatCell": {"range": rng(i), "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
                             "fields": "userEnteredFormat.numberFormat"}} for i in text_cols]
    reqs += [dim(i, {"pixelSize": w}, "pixelSize") for i, w in enumerate(widths)]
    if role:
        rule = lambda i, text, color: {"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [rng(i)], "booleanRule": {
            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": text}]}, "format": {"backgroundColor": rgb(color)}}}}}
        reqs += [
            {"setDataValidation": {"range": rng(0), "rule": {"condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": v} for v in STATUSES]},
                                                             "showCustomUi": True, "strict": True}}},
            rule(0, "Applied", "C6EFCE"), rule(0, "Pending", "FFEB9C"), rule(0, "Skip", "D9D9D9"), rule(1, "NEW", "BDD7EE"),
            dim(ncols - 1, {"hiddenByUser": True}, "hiddenByUser"),
            {"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"frozenColumnCount": 3}}, "fields": "gridProperties.frozenColumnCount"}},
        ]
    sh.batch_update({"requests": reqs})
    return sid, 6000


def write_tab(sh, title, header, rows, widths, text_cols, role=False):
    sid, nrows = ensure_tab(sh, title, len(header), widths, text_cols, role)
    if len(rows) + 100 > nrows:
        sh.batch_update({"requests": [{"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"rowCount": len(rows) + 500}},
                                                                   "fields": "gridProperties.rowCount"}}]})
    sh.values_clear(A1(title, f"A2:{get_column_letter(len(header))}"))
    sh.values_update(A1(title, "A1"), params={"valueInputOption": "USER_ENTERED"}, body={"values": [header] + rows})


def read_tab(sh, title, titles):
    if title not in titles:
        return []
    return sh.values_get(A1(title, "A1:Z"), params={"valueRenderOption": "FORMULA"}).get("values", [])


def link(u, label):
    return f'=HYPERLINK("{u.replace(chr(34), "%22")}","{label}")' if u else ""


def link_url(v):
    m = re.match(r'=HYPERLINK\("([^"]*)"', v or "")
    return m.group(1) if m else (v if (v or "").startswith("http") else "")


def read_state(sh, titles):
    state = {}
    for t in ROLES:
        v = read_tab(sh, t, titles)
        if len(v) < 2 or "ID" not in v[0]:
            continue
        ix = {h: i for i, h in enumerate(v[0])}
        for row in v[1:]:
            g = lambda k: row[ix[k]] if k in ix and ix[k] < len(row) else ""
            if not g("ID"):
                continue
            state[g("ID")] = dict(
                id=g("ID"), cat=t, status=g("Status") or "Pending", notes=g("Notes"), company=g("Company"), role=g("Role"),
                location=g("Location"), posted=g("Date Posted"), apply=link_url(g("Apply")), simplify=link_url(g("Simplify")),
                site=link_url(g("Company Site")), salary=g("Salary"), flags=g("Flags"),
                sources={x for x in g("Source").split(", ") if x}, first_seen=g("First Seen"), last_seen=g("Last Seen"))
    return state


# ---------- main flow ----------
def scrape():
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
    if len(fresh) < 300:
        raise SystemExit(f"Only {len(fresh)} roles scraped - something is wrong, leaving the Sheet untouched.")
    return fresh, counts, infos


def run(sh, import_xlsx=None):
    fresh, counts, infos = scrape()
    titles = meta(sh)  # read the Sheet only AFTER scraping, so the read -> write window stays short
    state = read_state(sh, titles)
    if not state and import_xlsx and Path(import_xlsx).exists():
        state = load_state(Path(import_xlsx))[0]
        print(f"Imported {len(state)} rows (with Applied marks) from {import_xlsx}")
    log = [[(l + [""] * 10)[i] if i != 9 else link_url((l + [""] * 10)[9]) for i in range(10)] for l in read_tab(sh, "Update Log", titles)[1:] if l]

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
        age = (TODAY - dt.date.fromisoformat(o["last_seen"])).days if re.match(r"\d{4}-\d\d-\d\d", o["last_seen"] or "") else 0
        if o["status"] in ("Applied",) or age <= 30:
            o.update(listed="No", sources=sorted(x for x in o["sources"] if x))
            final[i] = o

    by_cat = {r: [] for r in ROLES}
    for r in final.values():
        by_cat[r["cat"] if r["cat"] in by_cat else "Other"].append(r)
    recent = (TODAY - dt.timedelta(days=2)).isoformat()
    for name in ROLES:
        rows = sorted(by_cat[name], key=lambda r: (r["posted"] or "", r["company"]), reverse=True)
        write_tab(sh, name, COLS, [[
            r["status"], "NEW" if r["first_seen"] >= recent else "", r["company"], r["role"], r["location"], r["posted"],
            link(r["apply"], "Apply"), link(r["simplify"], "Simplify"), link(r["site"], "Site"), r["salary"], r["flags"],
            ", ".join(r["sources"]), r["first_seen"], r["last_seen"], r["listed"], r["notes"], r["id"]] for r in rows],
            WIDTHS, TEXT_COLS, role=True)

    last = {}
    for row in reversed(log):  # newest first
        last.setdefault(row[1], row[5])
    run_ts = NOW.strftime("%Y-%m-%d %H:%M")
    for s in SOURCES:
        c = infos[s["key"]]
        if not c["sha"] or last.get(s["name"]) == c["sha"]:
            continue  # nothing new on GitHub since the last logged commit -> no repeated log rows
        log.append([run_ts, s["name"], s["repo"], s["branch"], c["date"], c["sha"], "Yes", counts[s["key"]], new_by_src[s["key"]], c["url"]])
    log = sorted(log, key=lambda r: str(r[0]), reverse=True)[:600]
    write_tab(sh, "Update Log", LOG_COLS, [r[:9] + [link(r[9], "commit")] for r in log], LOG_W, LOG_TEXT)
    write_tab(sh, "Sources", ["Name", "Link"],
              [[s["name"], link(f"https://github.com/{s['repo']}/blob/{s['branch']}/README.md", f"github.com/{s['repo']}")] for s in SOURCES]
              + [[n, link(u, u)] for n, u in EXTRA_LINKS], [320, 520], [0])
    print(f"OK: {len(final)} rows | " + " | ".join(f"{n}: {len(by_cat[n])}" for n in ROLES))
    return final


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"))
    ap.add_argument("--import-xlsx", default=None, help="one-time import of Applied marks from an old tracker workbook")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if a.dry_run:
        fresh, counts, _ = scrape()
        print(f"dry run: {len(fresh)} roles", counts)
        return
    import gspread
    gc = gspread.service_account_from_dict(json.loads(os.environ["GOOGLE_SA_JSON"]))
    run(gc.open_by_key(a.sheet_id), a.import_xlsx)


if __name__ == "__main__":
    main()
