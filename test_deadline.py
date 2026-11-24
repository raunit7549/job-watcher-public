"""Self-check: a page whose JS never yields must be skipped at the deadline, and the next URL still scanned.

Run: python test_deadline.py   (takes ~DEADLINE seconds)
"""
import asyncio, time

import job_watcher as jw

jw.DEADLINE = 20
companies = {
    "Hang": ["data:text/html,<script>setTimeout(() => { for (;;) {} }, 50)</script>"],
    "Ok": ["data:text/html,<a href='https://example.com/job'>Graduate Trainee - Pune</a>"],
}
found, failed = [], []
t = time.time()
asyncio.run(jw.scan_all(companies, found, failed))
assert [c for c, _ in failed] == ["Hang"], failed
assert [j["title"] for j in found] == ["Graduate Trainee - Pune"], found
assert time.time() - t < jw.DEADLINE + 30, time.time() - t
print(f"ok in {time.time() - t:.0f}s")
