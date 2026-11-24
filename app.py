import html, io, json, re, threading, uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from urllib.parse import urljoin, urlparse
import requests
from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

app = Flask(__name__)
FILES, RUNS = {}, {}
BROWSER_LOCK = threading.Lock()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"
HDR = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
MAX_DETAIL = 60  # max job detail pages fetched per company

def clean(t):
    t = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", t or "")
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(html.unescape(t)))).strip()

def getj(url):
    r = requests.get(url, headers=HDR, timeout=20)
    return r.json() if r.status_code == 200 else None

# ================= 1. Fast paths: official ATS APIs =================
def greenhouse(s):
    d = getj(f"https://boards-api.greenhouse.io/v1/boards/{s}/jobs?content=true")
    if not d or "jobs" not in d: return None
    return [dict(title=j["title"], location=(j.get("location") or {}).get("name", ""),
                 dept=(j.get("departments") or [{}])[0].get("name", ""), url=j["absolute_url"],
                 text=clean(j.get("content")), posted=(j.get("updated_at") or "")[:10]) for j in d["jobs"]]

def lever(s):
    d = getj(f"https://api.lever.co/v0/postings/{s}?mode=json")
    if not isinstance(d, list) or not d: return None
    return [dict(title=j["text"], location=(j.get("categories") or {}).get("location", ""),
                 dept=(j.get("categories") or {}).get("team", ""), url=j["hostedUrl"],
                 text=j.get("descriptionPlain", "") + " " + " ".join(
                     l.get("text", "") + " " + clean(l.get("content")) for l in j.get("lists", [])),
                 posted="") for j in d]

def ashby(s):
    d = getj(f"https://api.ashbyhq.com/posting-api/job-board/{s}")
    if not d or not d.get("jobs"): return None
    return [dict(title=j["title"], location=j.get("location", ""), dept=j.get("department", ""),
                 url=j.get("jobUrl", ""), text=j.get("descriptionPlain", ""),
                 posted=(j.get("publishedAt") or "")[:10]) for j in d["jobs"]]

ATS = {"greenhouse.io": greenhouse, "lever.co": lever, "ashbyhq.com": ashby}

def ats_from_url(url):
    p = urlparse(url if "//" in url else "https://" + url)
    parts = [x for x in p.path.split("/") if x]
    for host, fn in ATS.items():
        if host in p.netloc and parts:
            return host, fn, parts[0]
    return None

# ================= 2. Generic engine: any careers URL =================
TITLE_KEYS = ("title", "jobTitle", "job_title", "positionTitle", "postingTitle", "name", "text")
URL_KEYS = ("url", "applyUrl", "apply_url", "jobUrl", "job_url", "hostedUrl", "absolute_url", "externalPath",
            "detailUrl", "jobPostingUrl", "canonicalPositionUrl", "link")
ID_KEYS = ("id", "jobId", "job_id", "requisitionId", "reqId", "req_id", "jobReqId", "positionId")
DESC_KEYS = ("description", "jobDescription", "descriptionPlain", "content", "summary", "shortDescription",
             "responsibilities", "qualifications")
LOC_KEYS = ("location", "locationName", "locations", "city", "jobLocation")
POSTED_KEYS = ("postedDate", "datePosted", "publishedAt", "created_at", "updated_at", "postedOn")
JOBLINK = re.compile(r"/(job|jobs|position|positions|opening|openings|vacanc\w*|requisition|posting|postings|"
                     r"jobdetail\w*|job-detail\w*|careers?/.+?)([/\-_?]|$)|[?&](jobid|job_id|req|reqid|requisition\w*)=", re.I)
JUNK = re.compile(r"^(apply|apply now|search|search jobs|view all|see all|learn more|read more|next|previous|"
                  r"back|home|careers|jobs|login|sign in|share|save)\b", re.I)

def first(d, keys):
    for k in keys:
        v = d.get(k)
        if isinstance(v, (str, int)) and not isinstance(v, bool) and str(v).strip():
            return k, str(v).strip()
    return None, None

def locstr(v):
    if isinstance(v, str): return v
    if isinstance(v, list): return "; ".join(x for x in (locstr(i) for i in v) if x)[:150]
    if isinstance(v, dict):
        return ", ".join(str(x) for k, x in v.items()
                         if k in ("name", "city", "addressLocality", "addressRegion", "state", "country") and isinstance(x, str))
    return ""

def walk(o, out, base):
    """Find job-like objects anywhere inside a JSON structure."""
    if isinstance(o, list):
        for x in o: walk(x, out, base)
    elif isinstance(o, dict):
        _, title = first(o, TITLE_KEYS)
        uk, url = first(o, URL_KEYS)
        _, jid = first(o, ID_KEYS)
        if title and 6 <= len(title) <= 150 and re.search(r"[A-Za-z]", title) and \
           (url or (jid and any(k in o for k in DESC_KEYS + LOC_KEYS))):
            if url and uk == "externalPath" and "myworkdayjobs" in base: full = base.rstrip("/") + url
            elif url: full = urljoin(base, url)
            else: full = base
            loc = next((locstr(o[k]) for k in LOC_KEYS if k in o and locstr(o[k])), "")
            out.append(dict(title=title, location=loc, dept=first(o, ("department", "team", "category"))[1] or "",
                            url=full, posted=(first(o, POSTED_KEYS)[1] or "")[:10],
                            text=clean(" ".join(str(o[k]) for k in DESC_KEYS if isinstance(o.get(k), str)))))
        for v in o.values():
            if isinstance(v, (dict, list)): walk(v, out, base)

def ld_json(page_html):
    for m in re.finditer(r'(?is)<script[^>]+type=["\'](?:application/ld\+json|application/json)["\'][^>]*>(.*?)</script>', page_html):
        try: yield json.loads(m.group(1))
        except Exception: pass

def ld_postings(page_html):
    res = []
    def rec(o):
        if isinstance(o, list): [rec(x) for x in o]
        elif isinstance(o, dict):
            t = o.get("@type")
            if t == "JobPosting" or (isinstance(t, list) and "JobPosting" in t): res.append(o)
            for v in o.values():
                if isinstance(v, (dict, list)): rec(v)
    for d in ld_json(page_html): rec(d)
    return res

def static_candidates(page_html, base):
    out = []
    for p in ld_postings(page_html):
        out.append(dict(title=str(p.get("title", "")), location=locstr((p.get("jobLocation") or {}).get("address", "")
                        if isinstance(p.get("jobLocation"), dict) else ""), dept="",
                        url=p.get("url") or base, posted=str(p.get("datePosted", ""))[:10], text=clean(p.get("description", ""))))
    for d in ld_json(page_html): walk(d, out, base)  # __NEXT_DATA__ and similar embedded state
    for m in re.finditer(r'(?is)<a\s[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page_html):
        href, text = m.group(1), clean(m.group(2))
        if href.startswith(("#", "mailto:", "javascript:")) or not (6 <= len(text) <= 140) or JUNK.match(text): continue
        full = urljoin(base, href)
        if JOBLINK.search(urlparse(full).path + "?" + urlparse(full).query):
            out.append(dict(title=text, location="", dept="", url=full, posted="", text=""))
    return out

def browser_candidates(url, log):
    """Render the page in headless Chromium, capture JSON API responses, scroll/paginate."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log("Browser engine not installed; used static HTML only"); return [], "", url
    found, content, final, seen_json = [], "", url, []
    with BROWSER_LOCK, sync_playwright() as p:
        b = p.chromium.launch(args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                                    "--disable-blink-features=AutomationControlled"])
        ctx = b.new_context(user_agent=UA, viewport={"width": 1366, "height": 900}, locale="en-US")
        ctx.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
        page = ctx.new_page()
        def on_resp(r):
            try:
                if "json" in r.headers.get("content-type", "").lower():
                    before = len(found)
                    walk(r.json(), found, page.url or url)
                    seen_json.append((r.url[:90], len(found) - before))
            except Exception: pass
        page.on("response", on_resp)
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=45000)
            log(f"Browser: HTTP {resp.status if resp else '?'}")
        except Exception as e:
            log(f"Browser load issue: {str(e)[:100]}")
        try: page.wait_for_load_state("networkidle", timeout=20000)
        except Exception: pass
        page.wait_for_timeout(4000)
        for _ in range(8):  # scroll + click "load more / next" a few times
            try:
                page.mouse.wheel(0, 8000); page.wait_for_timeout(1200)
                nxt = page.query_selector("button:has-text('Load more'), button:has-text('Show more'), "
                                          "a:has-text('Next'), button[aria-label*='next' i], a[aria-label*='next' i]")
                if nxt and nxt.is_visible() and nxt.is_enabled(): nxt.click(); page.wait_for_timeout(1800)
            except Exception: break
        try:
            content, final = page.content(), page.url
            body = clean(page.inner_text("body"))[:160] if page.query_selector("body") else ""
            log(f"Browser page: title='{page.title()[:60]}' text='{body}'")
        except Exception: pass
        log(f"Browser: {len(content)} chars HTML, {len(seen_json)} JSON responses, {len(found)} job-like items")
        for u, n in seen_json[:6]: log(f"  JSON {n:>3} jobs <- {u}")
        b.close()
    return found, content, final

def enrich(jobs, skills, log):
    """Fetch job detail pages for jobs lacking a description."""
    need = [j for j in jobs if len(j["text"]) < 200 and j["url"]]
    need.sort(key=lambda j: -len(skills_found(j["title"], skills)))
    def one(j):
        try:
            r = requests.get(j["url"], headers=HDR, timeout=15)
            if r.status_code != 200: return
            ld = ld_postings(r.text)
            new = clean(ld[0].get("description", "")) if ld and ld[0].get("description") else clean(r.text)[:20000]
            if len(new) > len(j["text"]): j["text"] = new
        except Exception: pass
    if need:
        log(f"Reading {min(len(need), MAX_DETAIL)} job pages…")
        with ThreadPoolExecutor(8) as ex: list(ex.map(one, need[:MAX_DETAIL]))

def generic_fetch(url, skills, log):
    url = url if "//" in url else "https://" + url
    jobs = []
    try:
        r = requests.get(url, headers=HDR, timeout=25)
        jobs = static_candidates(r.text, r.url) if r.status_code == 200 else []
        log(f"Static fetch: HTTP {r.status_code}, {len(r.text)} chars, {len(jobs)} candidates")
    except Exception as e:
        log(f"Static fetch failed: {e}")
    if len(jobs) < 5:  # JS-rendered site -> real browser
        log("Page looks JavaScript-rendered; opening in browser…")
        try: found, content, final = browser_candidates(url, log)
        except Exception as e: found, content, final = [], "", url; log(f"Browser unavailable: {str(e)[:120]}")
        jobs += found + (static_candidates(content, final) if content else [])
    seen, uniq = set(), []
    for j in jobs:
        k = (j["url"], j["title"].lower())
        if j["title"] and k not in seen: seen.add(k); uniq.append(j)
    enrich(uniq, skills, log)
    return uniq

def find_careers_url(company, log):
    s = re.sub(r"[^a-z0-9]", "", company.lower())
    for u in (f"https://careers.{s}.com", f"https://www.{s}.com/careers", f"https://{s}.com/careers",
              f"https://jobs.{s}.com", f"https://www.{s}.com/en/careers"):
        try:
            r = requests.get(u, headers=HDR, timeout=10)
            if r.status_code == 200 and re.search(r"career|jobs", r.text, re.I):
                log(f"Found careers page: {r.url}"); return r.url
        except Exception: pass
    return None

def fetch_jobs(company, url, skills, log):
    """Returns (method, jobs, note)."""
    hit = ats_from_url(url) if url else None
    cands = [(hit[1], hit[2])] if hit else []
    if not url:
        s = re.sub(r"[^a-z0-9]", "", company.lower())
        cands = [(fn, x) for x in dict.fromkeys([s, re.sub(r"[^a-z0-9]+", "-", company.lower()).strip("-")]) for fn in ATS.values()]
    for fn, slug in cands:
        try: jobs = fn(slug)
        except Exception: jobs = None
        if jobs: return f"ATS API ({slug})", jobs, ""
    if not url:
        url = find_careers_url(company, log)
        if not url: return None, None, "Could not locate careers page — paste its URL as: Company | https://..."
    jobs = generic_fetch(url, skills, log)
    return ("Generic scan", jobs, "") if jobs else (None, None, "No job listings detected on that page (login wall / bot protection?)")

# ================= 3. Matching =================
EXP = re.compile(r"(\d{1,2})\s*(?:\+|-|–|to)?\s*(\d{1,2})?\s*\+?\s*(?:years|yrs|year)", re.I)

def exp_range(text):
    m = EXP.search(text)
    return (int(m.group(1)), int(m.group(2)) if m.group(2) else None) if m else (None, None)

def skills_found(text, skills):
    low = text.lower()
    return [s for s in skills if re.search(r"(?<![\w+#.])" + re.escape(s.lower()) + r"(?![\w+#])", low)]

def search(companies, skills, years, tol, log):
    rows, report = [], []
    for company, url in companies:
        log(f"▶ {company}")
        try: method, jobs, note = fetch_jobs(company, url, skills, log)
        except Exception as e: method, jobs, note = None, None, f"Error: {e}"
        if not jobs:
            report.append((company, note or "No jobs found", 0, 0)); log(f"✗ {company}: {note}"); continue
        n = 0
        for j in jobs:
            blob = j["title"] + " " + j["text"]
            hit = skills_found(blob, skills)
            if not hit: continue
            lo, hi = exp_range(blob)
            if lo is not None and (lo > years + tol or (hi and years > hi + tol)): continue
            rows.append(dict(company=company, title=j["title"], location=j["location"], dept=j["dept"],
                             exp="Not specified" if lo is None else (f"{lo}-{hi} yrs" if hi else f"{lo}+ yrs"),
                             matched=", ".join(hit), pct=round(len(hit) / len(skills) * 100), posted=j["posted"],
                             source=method.split(" (")[0], url=j["url"]))
            n += 1
        report.append((company, f"OK - {method}", len(jobs), n)); log(f"✓ {company}: {n} matches of {len(jobs)} jobs")
    rows.sort(key=lambda r: (-r["pct"], r["company"]))
    return rows, report

def build_excel(rows, report, skills, years):
    wb = Workbook(); ws = wb.active; ws.title = "Matching Jobs"
    cols = [("Company", "company", 18), ("Job Title", "title", 45), ("Location", "location", 24),
            ("Department", "dept", 20), ("Experience Required", "exp", 18), ("Matched Skills", "matched", 32),
            ("Match %", "pct", 10), ("Posted", "posted", 12), ("Source", "source", 14), ("Apply Link", "url", 14)]
    ws.append([c[0] for c in cols])
    for i, c in enumerate(cols, 1):
        cell = ws.cell(1, i); cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78"); ws.column_dimensions[cell.column_letter].width = c[2]
    for r in rows:
        ws.append([r[c[1]] if c[1] != "url" else "Apply →" for c in cols])
        cell = ws.cell(ws.max_row, len(cols)); cell.hyperlink = r["url"]; cell.font = Font(color="0563C1", underline="single")
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
    s = wb.create_sheet("Summary")
    s.append(["Skills searched", ", ".join(skills)]); s.append(["Your experience (yrs)", years])
    s.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")]); s.append([])
    s.append(["Company", "Status", "Jobs scanned", "Matching jobs"])
    for c in s[5]: c.font = Font(bold=True)
    for r in report: s.append(list(r))
    s.column_dimensions["A"].width = 26; s.column_dimensions["B"].width = 70
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ================= 4. Web app (background runs + polling) =================
def run(rid, companies, skills, years, tol):
    st = RUNS[rid]
    try:
        rows, report = search(companies, skills, years, tol, lambda m: st["log"].append(m))
        FILES[rid] = build_excel(rows, report, skills, years)
        st.update(rows=rows, report=report, token=rid)
    except Exception as e:
        st["error"] = str(e)
    st["done"] = True

@app.post("/search")
def start():
    f = request.json; companies = []
    for line in f["companies"].splitlines():
        if line.strip():
            name, _, url = line.partition("|"); companies.append((name.strip(), url.strip()))
    skills = [s.strip() for s in f["skills"].split(",") if s.strip()]
    rid = uuid.uuid4().hex; RUNS[rid] = dict(log=[], done=False)
    threading.Thread(target=run, args=(rid, companies, skills, float(f["years"] or 0), 1.0), daemon=True).start()
    return jsonify(id=rid)

@app.get("/status/<rid>")
def status(rid): return jsonify(RUNS.get(rid, dict(done=True, error="Unknown run (server restarted?)", log=[])))

@app.get("/download/<token>")
def download(token):
    return send_file(io.BytesIO(FILES[token]), as_attachment=True, download_name="matching_jobs.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

PAGE = """<!doctype html><meta name=viewport content="width=device-width,initial-scale=1"><title>Job Finder</title>
<style>body{font-family:system-ui;max-width:900px;margin:30px auto;padding:0 16px}label{font-weight:600;display:block;margin-top:14px}
textarea,input{width:100%;padding:8px;box-sizing:border-box;font:inherit}button{margin-top:16px;padding:10px 18px;background:#1F4E78;color:#fff;border:0;border-radius:6px;font:inherit}
table{border-collapse:collapse;width:100%;margin-top:16px;font-size:14px}td,th{border:1px solid #ccc;padding:6px;text-align:left}small{color:#666}
pre{background:#f3f4f6;padding:8px;border-radius:6px;white-space:pre-wrap;font-size:12px;max-height:180px;overflow:auto}</style>
<h2>Tech Job Finder</h2>
<label>Companies (one per line)</label><textarea id=c rows=4 placeholder="Infosys | https://career.infosys.com/joblist&#10;Stripe"></textarea>
<small>Format: Company | careers page URL. URL optional — works with any careers site.</small>
<label>Tech stack (comma separated)</label><input id=s placeholder="Java, Spring, AWS">
<label>Your years of experience</label><input id=y type=number value=3 min=0>
<button id=b onclick=go()>Find jobs</button> <span id=st></span><pre id=lg hidden></pre><div id=out></div>
<script>
const e=t=>String(t).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
async function go(){b.disabled=true;out.innerHTML='';lg.hidden=false;lg.textContent='Starting…';st.textContent='Working…';
const r=await fetch('/search',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({companies:c.value,skills:s.value,years:y.value})});
const {id}=await r.json();
while(true){await new Promise(z=>setTimeout(z,2000));const d=await (await fetch('/status/'+id)).json();
lg.textContent=d.log.join('\\n');lg.scrollTop=lg.scrollHeight;
if(d.done){b.disabled=false;if(d.error){st.textContent='Error: '+d.error;return}
st.textContent=d.rows.length+' matching jobs';
out.innerHTML='<p><a href="/download/'+d.token+'"><b>⬇ Download Excel</b></a></p><table><tr><th>Company<th>Title<th>Location<th>Exp<th>Match<th>Link</tr>'+
d.rows.map(j=>`<tr><td>${e(j.company)}<td>${e(j.title)}<td>${e(j.location)}<td>${e(j.exp)}<td>${j.pct}%<td><a href="${e(j.url)}" target=_blank>Open</a></tr>`).join('')+
'</table><h4>Company status</h4><ul>'+d.report.map(x=>`<li>${e(x[0])}: ${e(x[1])} (${x[3]}/${x[2]})`).join('')+'</ul>';return}}}
</script>"""

@app.get("/")
def home(): return render_template_string(PAGE)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
