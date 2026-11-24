"""Daily job watcher: scans careers pages from the company Excel, emails new matches via Outlook.

Usage: python job_watcher.py            (normal run, sends email, updates state)
       python job_watcher.py --dry-run  (prints matches only)
       python job_watcher.py --profile business   (business/analytics roles; see PROFILES)
       python job_watcher.py --flush  (send a held geo section alone if the business run never merged it)
One mail a day: the geo run holds its section in HELD; the business run sends it to ME combined with its own.
"""
import asyncio, datetime, hashlib, html, json, os, re, sys
from pathlib import Path

from playwright.async_api import async_playwright

EXCEL = os.environ.get("COMPANY_EXCEL", "CompanyList.xlsx")  # only read with --refresh
ME = os.environ.get("MAIL_TO", "")            # gets the combined daily mail
OTHERS = os.environ.get("MAIL_BCC", "")       # comma-separated; get the business list only
GMAIL_USER = os.environ.get("GMAIL_USER", "")  # account whose app password is in GMAIL_APP_PASSWORD
PROFILES = {
    "geo": {
        "companies": "companies.json", "state": "seen_jobs.json", "label": "Job watch",
        "to": ME, "hold": True,
        "keywords": r"geolog|geophys|geoscien|\bGIS\b|remote sens|geospatial|geoinformat|survey|reservoir|petroleum|"
                    r"seismic|petrophysic|geochem|geotherm|energy analyst|research analyst|graduate|trainee|fresher|\bintern(ship)?s?\b|"
                    r"entry[- ]level|early career|"
                    r"environmental (scien|engineer|consult|impact|compliance|specialist|analyst)|\bEHS\b|\bHSE\b|"
                    r"sustainab|\bESG\b|ecolog|climate (risk|analyst|scien)|\bEIA\b|environmental (assessment|monitoring)",
        "exclude": r"senior|\blead\b|manager|principal|director|head of|\b([5-9]|1\d)\+? ?(yrs|years)",
        "terms": ("geo", "GIS", "graduate", "intern", "survey", "trainee", "environmental", "sustainability"),
    },
    "business": {
        "companies": "companies_business.json", "state": "seen_business.json", "label": "Business job watch",
        "to": ME, "bcc": OTHERS,
        "keywords": r"product (manager|intern|analyst|associate|management)|\bAPM\b|associate product|"
                    r"financial analyst|finance (analyst|intern|associate|graduate)|equity research|investment bank|"
                    r"\bFP&A\b|treasury analyst|credit analyst|risk analyst|valuation|corporate finance|"
                    r"consultant|consulting|business analyst|strategy (analyst|associate|intern|consultant)|"
                    r"data analyst|data scien|analytics (analyst|associate|intern|consultant|engineer)|"
                    r"business intelligence|\bBI (analyst|developer)|quantitative (analyst|research)|decision scien|"
                    r"graduate (program|analyst|trainee|scheme)|\btrainee\b|fresher|\bintern(ship)?s?\b|"
                    r"entry[- ]level|early career|campus hire|"
                    r"\banalyst\s*[-–,]|[-–,]\s*analyst\b",  # bank-style titles: "AWM - F&S - Analyst", "Analyst - Portfolio Risk"
        # no bare "manager": Product Manager is a target role here
        "exclude": r"senior|\blead\b|principal|director|head of|\bVP\b|vice president|senior manager|"
                   r"\b([5-9]|1\d)\+? ?(yrs|years)",
        "terms": ("intern", "graduate", "analyst", "product", "data", "consultant", "finance"),
    },
}
PROFILE = PROFILES[sys.argv[sys.argv.index("--profile") + 1] if "--profile" in sys.argv else "geo"]
KEYWORDS = re.compile(PROFILE["keywords"], re.I)
EXCLUDE = re.compile(PROFILE["exclude"], re.I)
SKIP_URL = re.compile(r"linkedin\.com/in/|zoominfo\.com/p/|theorg\.com", re.I)
# posting age/time some boards put in the title ("14 days ago", "07:43"); left in, it changes the key daily
STALE = re.compile(r"\b(?:(?:re)?posted\s+)?(?:\d+\+?|an?)\s+(?:minute|hour|day|week|month)s?\s+ago\b"
                   r"|\b(?:(?:re)?posted\s+)?(?:today|yesterday|just now)\b|\b\d{1,2}:\d{2}\b", re.I)
# shortcut: title/link only (no job-description fetch), so unmarked jobs may still take freshers
FRESHER = re.compile(r"fresher|graduate|\bgrad\b|trainee|entry[- ]level|early[- ]career|campus|"
                     r"\b0\s*(?:-|–|to)\s*[0-3]\s*\+?\s*(?:yrs?|years?)\b|\b0-to-[0-3]-years", re.I)
INTERN = re.compile(r"\bintern(ship)?s?\b", re.I)  # internships never get @f, even "Graduate Intern"
INDIA = re.compile(r"India|Mumbai|Pune|Bengaluru|Bangalore|Delhi|Gurugram|Gurgaon|Noida|Hyderabad|Chennai|Kolkata|remote", re.I)

HERE = Path(__file__).parent
STATE = HERE / PROFILE["state"]
HELD = HERE / "held_mail.html"  # geo section waiting for the business run (first line = subject)
LOG = HERE / "run.log"
DEADLINE = 240  # seconds per URL: one stuck page/socket must not eat the whole job


def log(msg):
    line = f"{datetime.datetime.now():%Y-%m-%d %H:%M} {msg}"
    print(line)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_companies():
    """companies.json (names + URLs only) if present, else read the Excel and write that json."""
    cache = HERE / PROFILE["companies"]
    if cache.exists() and "--refresh" not in sys.argv:
        return json.loads(cache.read_text(encoding="utf-8"))
    import pandas as pd
    df = pd.read_excel(EXCEL, sheet_name="Company list")
    out = {}
    for _, r in df[df["Company Name"].notna()].iterrows():
        name = str(r["Company Name"]).strip().strip("\u2019\ufffd ")
        urls = [u.strip() for c in (r.iloc[9], r.iloc[6]) if pd.notna(c)
                for u in re.split(r"[\s|,]+", str(c)) if u.startswith("http")]
        if urls:  # column J first, website G as fallback
            out.setdefault(name, [])
            for u in urls[:2]:
                if u not in out[name]:
                    out[name].append(u)
    cache.write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8")
    return out


def is_match(text):
    return 8 <= len(text) <= 160 and KEYWORDS.search(text) and not EXCLUDE.search(text)


def mk(company, title, link):
    title = " ".join(STALE.sub(" ", title).split())
    return {"company": company, "title": title, "link": link, "india": bool(INDIA.search(title + link)),
            "fresher": bool(FRESHER.search(f"{title} {link}")) and not INTERN.search(f"{title} {link}"),
            "id": hashlib.sha1(f"{company}|{title}|{link}".encode()).hexdigest(),
            # link-independent key: same company+title is one entry even if the URL/tracking params change
            "key": hashlib.sha1(f"{company}|{title}".lower().encode()).hexdigest()}


def api_jobs(company, url):
    """Ask the ATS for its job list directly. Returns None if the URL isn't a known ATS."""
    import requests
    g = requests.get
    try:
        m = re.match(r"https://([\w-]+)\.(wd\d+)\.myworkdayjobs\.com/(?:\w\w-\w\w/)?([\w-]+)", url)
        if m:  # Workday: undocumented but stable JSON search endpoint the portal itself calls
            sub, wd, site = m.groups()
            out = {}
            for term in PROFILE["terms"]:
                r = requests.post(f"https://{sub}.{wd}.myworkdayjobs.com/wday/cxs/{sub}/{site}/jobs",
                                  json={"appliedFacets": {}, "limit": 20, "offset": 0, "searchText": term},
                                  headers={"Accept": "application/json", "User-Agent": "Mozilla/5.0"}, timeout=30)
                for j in r.json().get("jobPostings", []):
                    t = f"{j['title']} - {j.get('locationsText', '')}".strip(" -")
                    out[t] = f"https://{sub}.{wd}.myworkdayjobs.com/{site}{j['externalPath']}"
            return [mk(company, t, h) for t, h in out.items()]
        if "jobs.exxonmobil.com" in url:  # SuccessFactors: server-rendered India search, 25 rows a page
            from bs4 import BeautifulSoup
            out = {}
            for term in PROFILE["terms"]:
                for row in range(0, 50, 25):
                    r = g("https://jobs.exxonmobil.com/search/", params={"q": term, "locationsearch": "India", "startrow": row},
                          headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
                    for tr in BeautifulSoup(r.text, "html.parser").select("tr.data-row"):
                        a = tr.select_one("a.jobTitle-link")
                        loc = tr.select_one(".jobLocation")
                        if a:
                            out[f"{a.get_text(strip=True)} - {loc.get_text(strip=True) if loc else ''}".strip(" -")] =                                 "https://jobs.exxonmobil.com" + a["href"]
            return [mk(company, t, h) for t, h in out.items()]
        if "careers.slb.com" in url:  # Coveo search behind the page; its public token is embedded in the HTML, so read it fresh
            html_ = g("https://careers.slb.com/job-listing", headers={"User-Agent": "Mozilla/5.0"}, timeout=30).text
            org = re.search(r'organizationId" value="(\w+)"', html_).group(1)
            tok = re.search(r'accessToken" value="([^"]+)"', html_).group(1)
            r = requests.post(f"https://{org}.org.coveo.com/rest/search/v2?organizationId={org}",
                              headers={"Authorization": f"Bearer {tok}"}, timeout=30,
                              json={"q": "", "numberOfResults": 200, "pipeline": "ATSJobsPipeline", "searchHub": "CoveoJobsHub",
                                    "cq": '@source=="ATS_Jobs_Source - Prod" @country=="India"'})  # all India postings (~25), filtered by is_match
            out = {f"{j['title']} - India": j["clickUri"] for j in r.json()["results"]}
            return [mk(company, t, h) for t, h in out.items()]
        m = re.match(r"https://([\w-]+)\.bamboohr\.com", url)
        if m:
            r = g(f"https://{m.group(1)}.bamboohr.com/careers/list", timeout=30).json()
            return [mk(company, f"{j['jobOpeningName']} - {(j.get('location') or {}).get('city', '')}",
                       f"https://{m.group(1)}.bamboohr.com/careers/{j['id']}") for j in r.get("result", [])]
        m = re.match(r"https://apply\.workable\.com/([\w-]+)", url)
        if m:
            r = g(f"https://apply.workable.com/api/v1/widget/accounts/{m.group(1)}?details=true", timeout=30).json()
            return [mk(company, f"{j['title']} - {j.get('city', '')}", j["url"]) for j in r.get("jobs", [])]
        m = re.match(r"https://boards\.greenhouse\.io/([\w-]+)", url)
        if m:
            r = g(f"https://boards-api.greenhouse.io/v1/boards/{m.group(1)}/jobs", timeout=30).json()
            return [mk(company, f"{j['title']} - {j['location']['name']}", j["absolute_url"]) for j in r.get("jobs", [])]
        m = re.match(r"https://([\w.-]+\.oraclecloud\.com)/hcmUI/CandidateExperience/\w+/sites/(\w+)", url)
        if m:  # Oracle Recruiting Cloud (AmEx, Goldman Sachs, Verisk): public REST feed behind the candidate site
            host, site = m.groups()
            out = {}
            for term in PROFILE["terms"]:
                r = g(f"https://{host}/hcmRestApi/resources/latest/recruitingCEJobRequisitions?onlyData=true"
                      f"&expand=requisitionList&finder=findReqs;siteNumber={site},limit=50,keyword={term},"
                      f"locationId=300000000228786,sortBy=POSTING_DATES_DESC",  # Oracle's shared geography id for India
                      timeout=30).json()
                for j in r["items"][0]["requisitionList"]:
                    out[f"{j['Title']} - {j.get('PrimaryLocation', '')}"] = \
                        f"https://{host}/hcmUI/CandidateExperience/en/sites/{site}/job/{j['Id']}"
            return [mk(company, t, h) for t, h in out.items()]
        m = re.match(r"https://([\w.-]+)/careers\?domain=([\w.]+)", url)
        if m:  # Eightfold (HSBC): India-filtered search API the portal itself calls
            host, dom = m.groups()
            out = {}
            for term in PROFILE["terms"]:
                r = g(f"https://{host}/api/apply/v2/jobs?domain={dom}&query={term}&location=India&num=50",
                      headers={"User-Agent": "Mozilla/5.0"}, timeout=30).json()
                for j in r.get("positions", []):
                    out[f"{j['name']} - {j.get('location', '')}"] = j["canonicalPositionUrl"]
            return [mk(company, t, h) for t, h in out.items()]
        m = re.match(r"https://(?:jobs|careers)\.smartrecruiters\.com/([\w-]+)", url)
        if m:  # India postings only: global boards run to hundreds and the API pages at 100
            r = g(f"https://api.smartrecruiters.com/v1/companies/{m.group(1)}/postings?limit=100&country=in", timeout=30).json()
            return [mk(company, f"{j['name']} - {j['location'].get('city', '')}",
                       f"https://jobs.smartrecruiters.com/{m.group(1)}/{j['id']}") for j in r.get("content", [])]
        m = re.match(r"https://jobs\.lever\.co/([\w-]+)", url)
        if m:
            r = g(f"https://api.lever.co/v0/postings/{m.group(1)}?mode=json", timeout=30).json()
            return [mk(company, f"{j['text']} - {j['categories'].get('location', '')}", j["hostedUrl"]) for j in r]
    except Exception as ex:
        log(f"{company}: API attempt failed ({str(ex)[:80]})")
    return None


def plain_scan(company, url):
    """No browser: some hosts block Chromium but answer a plain request fine."""
    import requests
    from bs4 import BeautifulSoup
    r = requests.get(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/128"},
                     timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    jobs = {}
    for a in soup.find_all("a", href=True):
        text = " ".join(a.get_text().split())
        if is_match(text):
            jobs[text] = requests.compat.urljoin(url, a["href"])
    if not jobs:
        for text in (" ".join(l.split()) for l in soup.get_text("\n").splitlines()):
            if is_match(text) and len(text) <= 100:
                jobs[text] = url
    return [mk(company, t, h) for t, h in jobs.items()]


async def scan(page, company, url):
    await page.goto(url, timeout=45000, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass  # heavy portals never go idle; use what loaded
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")  # trigger lazy-loaded job lists
    await page.wait_for_timeout(1200)
    if not (await page.inner_text("body")).strip():  # bot walls (Incapsula) serve a blank page: fail loudly, not "0 matches"
        raise RuntimeError("blank page, likely blocked")
    jobs = {}
    for a in await page.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText, e.href])"):
        text = " ".join(a[0].split())
        if is_match(text) and a[1].startswith("http") and "theorg.com" not in a[1]:
            jobs[text] = a[1]
    if not jobs:  # no job links: fall back to plain text lines (static career pages)
        for line in (await page.inner_text("body")).splitlines():
            text = " ".join(line.split())
            if is_match(text) and len(text) <= 100:
                jobs[text] = url
    return [mk(company, t, h) for t, h in jobs.items()]


def send_mail(subject, body, to=None, bcc=None):
    body = "<p><b>@f</b> = for freshers (title or link says fresher / graduate / trainee / entry-level / 0-2 yrs)</p>" + body
    to = PROFILE.get("to", ME) if to is None else to
    bcc = PROFILE.get("bcc", "") if bcc is None else bcc
    pw = os.environ.get("GMAIL_APP_PASSWORD")
    if pw:  # cloud (GitHub Actions): Gmail SMTP, sends to itself
        import smtplib
        from email.mime.text import MIMEText
        msg = MIMEText(body, "html")
        msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, to  # no Bcc header: recipients stay hidden
        all_recipients = [a.strip() for a in (to + ("," + bcc if bcc else "")).split(",") if a.strip()]
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
            s.login(GMAIL_USER, pw.replace(" ", ""))
            s.send_message(msg, to_addrs=all_recipients)
        return
    import win32com.client
    mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
    mail.To, mail.Subject, mail.HTMLBody = to.replace(",", ";"), subject, body  # Outlook wants semicolons
    if bcc:
        mail.BCC = bcc.replace(",", ";")
    mail.Send()


def build_body(new, failed, baseline):
    e = html.escape
    parts = ["<p><b>Baseline run</b>: everything currently listed. Future emails show only new items.</p>" if baseline
             else f"<p>{len(new)} new matching item(s).</p>"]
    by_co = {}
    for j in new:
        by_co.setdefault(j["company"], []).append(j)
    for co, jobs in sorted(by_co.items()):
        rows = "".join(f"<tr><td>{'<b>@f</b> ' if j.get('fresher') else ''}<a href='{e(j['link'])}'>{e(j['title'])}</a></td>"
                       f"<td>{'India' if j['india'] else ''}</td></tr>" for j in jobs)
        parts.append(f"<h3>{e(co)}</h3><table border='1' cellpadding='4' style='border-collapse:collapse'>{rows}</table>")
    if failed:
        parts.append("<h3>Could not read</h3><ul>" + "".join(f"<li>{e(c)}: {e(u)}</li>" for c, u in failed) + "</ul>")
    return "".join(parts)


async def _scan_url(ctx, co, url, found, failed):
    jobs = await asyncio.to_thread(api_jobs, co, url)  # thread: blocking requests must stay cancellable by the deadline
    if jobs is not None:
        found += [j for j in jobs if is_match(j["title"])]
        log(f"{co}: {len(jobs)} job(s) from ATS API at {url}")
        return
    for attempt in (1, 2):  # one retry: most failures here are slow first loads
        page = await ctx.new_page()  # fresh tab per URL so one crash can't poison the rest
        try:
            jobs = await scan(page, co, url)
            found += jobs
            log(f"{co}: {len(jobs)} match(es) at {url}")
            break
        except Exception as ex:
            err = str(ex).splitlines()[0][:120]
            if attempt == 2 or "NAME_NOT_RESOLVED" in err:
                try:  # last resort: no browser
                    jobs = await asyncio.to_thread(plain_scan, co, url)
                    found += jobs
                    log(f"{co}: {len(jobs)} match(es) at {url} (plain HTTP)")
                except Exception as ex2:
                    failed.append((co, url))
                    log(f"{co}: FAILED {url} ({err} | plain: {str(ex2)[:60]})")
                break  # fallback done: retrying would log/count the same URL twice
        finally:
            await page.close()


async def scan_all(companies, found, failed):
    # async API on purpose: wait_for cancels a hung Playwright call cleanly. The old SIGALRM deadline raised inside
    # the sync API's dispatcher greenlet, killing it, so the next call (page.close) spun forever and the run timed out.
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="msedge" if os.name == "nt" else None)  # Edge locally, Chromium in cloud
        ctx = await browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/128 Safari/537.36")
        for co, urls in companies.items():
            for url in urls:
                if SKIP_URL.search(url):  # personal profiles: never list jobs
                    continue
                try:
                    await asyncio.wait_for(_scan_url(ctx, co, url, found, failed), DEADLINE)
                except TimeoutError:
                    failed.append((co, url))
                    log(f"{co}: STUCK >{DEADLINE}s, skipped {url}")
        await browser.close()


def main():
    if "--flush" in sys.argv:  # workflow fallback: business run died before merging, so send the geo section alone
        if HELD.exists():
            subject, body = HELD.read_text(encoding="utf-8").split("\n", 1)
            send_mail(subject, body, to=ME, bcc="")
            HELD.unlink()
            log(f"Sent held '{subject}'")
        return
    dry = "--dry-run" in sys.argv
    companies = load_companies()
    found, failed = [], []
    asyncio.run(scan_all(companies, found, failed))

    if dry:
        for j in found:
            print(f"  [{j['company']}] {j['title']}  ->  {j['link']}")
        return

    baseline = not STATE.exists()
    seen = set() if baseline else set(json.loads(STATE.read_text()))
    # dedupe within the run by key; legacy state holds ids, so accept either id or key as "seen"
    new = [j for j in {j["key"]: j for j in found}.values() if j["key"] not in seen and j["id"] not in seen]
    STATE.write_text(json.dumps(sorted(seen | {j["id"] for j in found} | {j["key"] for j in found})))

    today = datetime.date.today().strftime("%d %b %Y")
    subject = f"{PROFILE['label']} {today}: " + ("baseline" if baseline else f"{len(new)} new")
    body = build_body(new, failed, baseline)
    if PROFILE.get("hold"):
        HELD.write_text(f"{subject}\n{body}", encoding="utf-8")
        log(f"Held '{subject}' for the combined mail")
        return
    if HELD.exists():  # one mail for ME with both sections; bcc list still gets the business list only
        geo_subject, geo_body = HELD.read_text(encoding="utf-8").split("\n", 1)
        if PROFILE.get("bcc"):
            send_mail(subject, body, to="")
        tail = lambda s: s.split(": ", 1)[-1]
        send_mail(f"Job watch {today}: {tail(geo_subject)} geo, {tail(subject)} business",
                  f"<h2>Geo / environment</h2>{geo_body}<h2>Business</h2>{body}", to=ME, bcc="")
        HELD.unlink()
        log(f"Sent combined mail + business-only bcc ({len(failed)} unreadable)")
        return
    send_mail(subject, body)
    log(f"Sent '{subject}' ({len(failed)} unreadable)")


if __name__ == "__main__":
    main()
