# Job Watcher

A daily job alert for campus job seekers. Every night it reads the careers pages of 364 companies, keeps only fresher-relevant roles, and emails each person the postings that are new since yesterday.

Running on GitHub Actions since 17 September 2026, for me and six batchmates at IIT Kanpur.

## The problem

Off-campus openings are spread across hundreds of company careers pages. Each company uses a different portal, most postings are for experienced hires, and the fresher ones close within days. Checking them by hand every day is not possible, and job aggregators miss or delay many of them.

## What it does

1. **Reads each careers page.** For ten applicant-tracking systems (Workday, Oracle Recruiting, Eightfold, SmartRecruiters, Greenhouse, Lever, Workable, BambooHR, SuccessFactors, Coveo) it asks the portal's own job feed directly. For everything else it opens the page in a headless browser, and falls back to a plain HTTP request when the browser is blocked.
2. **Filters by role.** Two profiles: `geo` (geoscience, GIS, environment, energy) and `business` (product, analytics, consulting, finance). Senior roles and roles asking 5+ years are dropped.
3. **Shows only new postings.** A posting is one company + title, so a changed link or a "14 days ago" label does not make it look new again.
4. **Sends one mail a day**, grouped by company, before 7 am IST. Fresher roles are marked `@f`. Pages that could not be read are listed at the bottom, so a silent failure is never mistaken for "no new jobs".

## Run it

```
pip install playwright requests beautifulsoup4 && playwright install chromium
python job_watcher.py --dry-run                      # print matches, send nothing
python job_watcher.py --profile business --dry-run
python test_job_watcher.py                           # self-checks
```

To send mail, set `GMAIL_USER`, `GMAIL_APP_PASSWORD`, `MAIL_TO` and `MAIL_BCC` (as repository secrets for the scheduled workflow in `.github/workflows/daily.yml`).

## Files

| File | What it is |
|---|---|
| `job_watcher.py` | The whole watcher |
| `companies.json`, `companies_business.json` | Watchlists: company name and careers URLs |
| `test_job_watcher.py`, `test_deadline.py` | Self-checks for de-duplication, the fresher mark, the one-mail-a-day flow and the per-page deadline |

This is a public copy of the code. The live run, its recipient list and its run history are in a private repository.
