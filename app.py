import html, io, re, uuid
from datetime import datetime
from urllib.parse import urlparse
import requests
from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

app = Flask(__name__)
FILES = {}  # token -> excel bytes
HDR = {"User-Agent": "Mozilla/5.0 JobFinder/1.0"}

def clean(t):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(html.unescape(t or "")))).strip()

def get(url):
    r = requests.get(url, headers=HDR, timeout=20)
    return r.json() if r.status_code == 200 else None

# ---- ATS fetchers: each returns a list of normalized jobs or None ----
def greenhouse(s):
    d = get(f"https://boards-api.greenhouse.io/v1/boards/{s}/jobs?content=true")
    if not d or "jobs" not in d: return None
    return [dict(title=j["title"], location=(j.get("location") or {}).get("name", ""),
                 dept=(j.get("departments") or [{}])[0].get("name", ""), url=j["absolute_url"],
                 text=clean(j.get("content")), posted=(j.get("updated_at") or "")[:10]) for j in d["jobs"]]

def lever(s):
    d = get(f"https://api.lever.co/v0/postings/{s}?mode=json")
    if not isinstance(d, list) or not d: return None
    return [dict(title=j["text"], location=(j.get("categories") or {}).get("location", ""),
                 dept=(j.get("categories") or {}).get("team", ""), url=j["hostedUrl"],
                 text=j.get("descriptionPlain", "") + " " + " ".join(
                     l.get("text", "") + " " + clean(l.get("content")) for l in j.get("lists", [])),
                 posted=datetime.fromtimestamp(j["createdAt"] / 1000).strftime("%Y-%m-%d") if j.get("createdAt") else "")
            for j in d]

def ashby(s):
    d = get(f"https://api.ashbyhq.com/posting-api/job-board/{s}")
    if not d or not d.get("jobs"): return None
    return [dict(title=j["title"], location=j.get("location", ""), dept=j.get("department", ""),
                 url=j.get("jobUrl", ""), text=j.get("descriptionPlain", ""),
                 posted=(j.get("publishedAt") or "")[:10]) for j in d["jobs"]]

ATS = {"Greenhouse": greenhouse, "Lever": lever, "Ashby": ashby}

def slugs(company, url):
    out = []
    if url:
        p = urlparse(url if "//" in url else "https://" + url)
        parts = [x for x in p.path.split("/") if x]
        if any(k in p.netloc for k in ("greenhouse", "lever", "ashbyhq")) and parts:
            out.append(parts[0])
        out.append(p.netloc.replace("www.", "").replace("careers.", "").replace("jobs.", "").split(".")[0])
    c = company.lower().strip()
    out += [re.sub(r"[^a-z0-9]", "", c), re.sub(r"[^a-z0-9]+", "-", c).strip("-")]
    return list(dict.fromkeys(x for x in out if x))

def fetch_jobs(company, url):
    for s in slugs(company, url):
        for name, fn in ATS.items():
            try:
                jobs = fn(s)
            except Exception:
                jobs = None
            if jobs: return name, jobs
    return None, None

# ---- matching ----
EXP = re.compile(r"(\d{1,2})\s*(?:\+|-|–|to)?\s*(\d{1,2})?\s*\+?\s*(?:years|yrs|year)", re.I)

def exp_range(text):
    m = EXP.search(text)
    if not m: return None, None
    lo, hi = int(m.group(1)), int(m.group(2)) if m.group(2) else None
    return lo, hi

def skills_found(text, skills):
    low = text.lower()
    return [s for s in skills if re.search(r"(?<![\w+#.])" + re.escape(s.lower()) + r"(?![\w+#])", low)]

def search(companies, skills, years, tol):
    rows, report = [], []
    for company, url in companies:
        ats, jobs = fetch_jobs(company, url)
        if not jobs:
            report.append((company, "Not found (custom/unsupported careers page — paste the exact careers URL)", 0, 0))
            continue
        n = 0
        for j in jobs:
            blob = j["title"] + " " + j["text"]
            hit = skills_found(blob, skills)
            if not hit: continue
            lo, hi = exp_range(blob)
            if lo is not None and (lo > years + tol or (hi and years > hi + tol)): continue
            req = "Not specified" if lo is None else (f"{lo}-{hi} yrs" if hi else f"{lo}+ yrs")
            rows.append(dict(company=company, title=j["title"], location=j["location"], dept=j["dept"],
                             exp=req, matched=", ".join(hit), pct=round(len(hit) / len(skills) * 100),
                             posted=j["posted"], source=ats, url=j["url"]))
            n += 1
        report.append((company, f"OK via {ats}", len(jobs), n))
    rows.sort(key=lambda r: (-r["pct"], r["company"]))
    return rows, report

def build_excel(rows, report, skills, years):
    wb = Workbook(); ws = wb.active; ws.title = "Matching Jobs"
    cols = [("Company", "company", 18), ("Job Title", "title", 45), ("Location", "location", 24),
            ("Department", "dept", 20), ("Experience Required", "exp", 18), ("Matched Skills", "matched", 32),
            ("Match %", "pct", 10), ("Posted", "posted", 12), ("Source", "source", 12), ("Apply Link", "url", 14)]
    ws.append([c[0] for c in cols])
    for i, c in enumerate(cols, 1):
        cell = ws.cell(1, i); cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        ws.column_dimensions[cell.column_letter].width = c[2]
    for r in rows:
        ws.append([r[c[1]] if c[1] != "url" else "Apply →" for c in cols])
        cell = ws.cell(ws.max_row, len(cols))
        cell.hyperlink = r["url"]; cell.font = Font(color="0563C1", underline="single")
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
    s = wb.create_sheet("Summary")
    s.append(["Skills searched", ", ".join(skills)]); s.append(["Your experience (yrs)", years])
    s.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")]); s.append([])
    s.append(["Company", "Status", "Total jobs scanned", "Matching jobs"])
    for c in s[5]: c.font = Font(bold=True)
    for r in report: s.append(list(r))
    s.column_dimensions["A"].width = 26; s.column_dimensions["B"].width = 70
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

@app.post("/search")
def do_search():
    f = request.json
    companies = []
    for line in f["companies"].splitlines():
        line = line.strip()
        if not line: continue
        name, _, url = line.partition("|")  # "Stripe | https://boards.greenhouse.io/stripe"
        companies.append((name.strip(), url.strip()))
    skills = [s.strip() for s in f["skills"].split(",") if s.strip()]
    years = float(f["years"] or 0); tol = float(f.get("tol", 1))
    rows, report = search(companies, skills, years, tol)
    token = uuid.uuid4().hex
    FILES[token] = build_excel(rows, report, skills, years)
    return jsonify(token=token, rows=rows, report=report)

@app.get("/download/<token>")
def download(token):
    return send_file(io.BytesIO(FILES[token]), as_attachment=True, download_name="matching_jobs.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

PAGE = """<!doctype html><meta name=viewport content="width=device-width,initial-scale=1"><title>Job Finder</title>
<style>body{font-family:system-ui;max-width:900px;margin:30px auto;padding:0 16px}label{font-weight:600;display:block;margin-top:14px}
textarea,input{width:100%;padding:8px;box-sizing:border-box;font:inherit}button{margin-top:16px;padding:10px 18px;background:#1F4E78;color:#fff;border:0;border-radius:6px;font:inherit;cursor:pointer}
table{border-collapse:collapse;width:100%;margin-top:16px;font-size:14px}td,th{border:1px solid #ccc;padding:6px;text-align:left}small{color:#666}</style>
<h2>Tech Job Finder</h2>
<label>Companies (one per line)</label><textarea id=c rows=4 placeholder="Stripe&#10;Figma&#10;Acme | https://boards.greenhouse.io/acme"></textarea>
<small>Optional: add “| careers URL” after a name if auto-detect fails.</small>
<label>Tech stack (comma separated)</label><input id=s placeholder="Python, Django, AWS, React">
<label>Your years of experience</label><input id=y type=number value=3 min=0>
<button onclick=go()>Find jobs</button> <span id=st></span><div id=out></div>
<script>
async function go(){st.textContent='Searching…';out.innerHTML='';
const r=await fetch('/search',{method:'POST',headers:{'Content-Type':'application/json'},
body:JSON.stringify({companies:c.value,skills:s.value,years:y.value})});const d=await r.json();
st.textContent=d.rows.length+' matching jobs';
const e=t=>String(t).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
out.innerHTML='<p><a href="/download/'+d.token+'"><b>⬇ Download Excel</b></a></p><table><tr><th>Company<th>Title<th>Location<th>Exp<th>Match<th>Link</tr>'+
d.rows.map(j=>`<tr><td>${e(j.company)}<td>${e(j.title)}<td>${e(j.location)}<td>${e(j.exp)}<td>${j.pct}%<td><a href="${e(j.url)}" target=_blank>Open</a></tr>`).join('')+
'</table><h4>Company status</h4><ul>'+d.report.map(x=>`<li>${e(x[0])}: ${e(x[1])} (${x[3]}/${x[2]})`).join('')+'</ul>'}
</script>"""

@app.get("/")
def home(): return render_template_string(PAGE)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
