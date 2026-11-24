"""Run: python test_job_watcher.py  — titles taken from the 10/11 Oct 2026 watch emails."""
import os, pathlib, tempfile
os.environ.update(MAIL_TO="me@example.com", MAIL_BCC="friend@example.com")
import job_watcher as jw
from job_watcher import mk

pairs = [  # same posting scraped on two different days must give one key
    ("Bengaluru Consumer Payments Business Business Operations Analyst, Payments Business Full-time 14 days ago India",
     "Bengaluru Consumer Payments Business Business Operations Analyst, Payments Business Full-time 15 days ago India"),
    ("Associate Business Analyst 5 days ago Bangalore 0-1 Full Time India",
     "Associate Business Analyst 6 days ago Bangalore 0-1 Full Time India"),
    ("Analyst - Portfolio and Treasury Services Mumbai, India 05:13 India",
     "Analyst - Portfolio and Treasury Services Mumbai, India 04:21 India"),
    ("FY26 Intern - SW Architecture - Hod Hasharon Engineering Intern Posted a day ago",
     "FY26 Intern - SW Architecture - Hod Hasharon Engineering Intern Posted 2 days ago"),
]
for a, b in pairs:
    assert mk("X", a, "u")["key"] == mk("X", b, "u")["key"], (a, b)
assert mk("X", "Data Analyst 21 days ago", "u")["title"] == "Data Analyst"
assert mk("X", "Graduate Trainee - 2027", "u")["title"] == "Graduate Trainee - 2027"  # real content untouched

# @f fresher mark: title or link signal only
for t, link in [("Product Manager, New Grad (2027) - Amsterdam", "u"), ("Data Analyst (Fresher)", "u"),
                ("Graduate Management Trainee", "u"), ("Data Analyst 0-2 Yrs Pune", "u"),
                ("Engineer Analyst", "https://www.naukri.com/job-listings-engineer-analyst-l-t-chennai-0-to-2-years-1")]:
    assert mk("X", t, link)["fresher"], t
for t in ["Senior Product Analyst", "Analyst - Portfolio and Treasury Services", "Manager 10-12 years", "Lead 0-5 years",
          "Product Manager 0-1 products",  # bare "0-1" is PM jargon (zero-to-one), not experience
          "Graduate Intern - Data", "Product Management Intern (2027) - Berlin", "Graduate Internship Programme"]:
    assert not mk("X", t, "u")["fresher"], t

# one mail a day: geo run holds, business run sends bcc list its own section and ME the combined mail
sent = []
jw.send_mail = lambda s, b, to=None, bcc=None: sent.append((s, to, bcc))
jw.load_companies, jw.log = (lambda: {}), (lambda m: None)
tmp = pathlib.Path(tempfile.mkdtemp())
jw.HELD = tmp / "held.html"
for name, title in (("geo", "Geologist"), ("business", "Data Analyst")):
    async def fake_scan(companies, found, failed, t=title):
        found.append(mk("Co", t, "u"))
    jw.scan_all, jw.PROFILE, jw.STATE = fake_scan, jw.PROFILES[name], tmp / f"{name}.json"
    jw.STATE.write_text("[]")
    jw.main()
assert not jw.HELD.exists()
assert [(to, bcc) for _, to, bcc in sent] == [("", None), (jw.ME, "")], sent
assert sent[1][0].endswith("1 new geo, 1 new business"), sent[1][0]
print("ok")
