# ============================================================
#  STEP 0 — INSTALL LIBRARIES (run once in your terminal)
#
#  pip install requests pandas
#
#  WHY requests (not Selenium)?
#  Your lecture says: "Use requests + BeautifulSoup for static pages.
#  Use Selenium when content loads via JavaScript."
#
#  jobs.krd is a JavaScript app, BUT after inspecting the network
#  traffic (browser DevTools → Network tab) we discovered it calls
#  a public REST API that returns plain JSON.  We can call that
#  API directly with requests — no browser required, and we get
#  cleaner, richer data.
#
#  This is the 5-step pipeline from the slides:
#   Step 1 → Inspect the page (done: found the API endpoint)
#   Step 2 → Check robots.txt (allows crawling)
#   Step 3 → Fetch pages with requests
#   Step 4 → Parse & extract (JSON, so no BeautifulSoup needed)
#   Step 5 → Save to CSV with pandas
# ============================================================

# ── STEP 1: IMPORT LIBRARIES ────────────────────────────────

import time
import requests
import pandas as pd

# ── STEP 2: CONFIGURE THE SCRAPER ───────────────────────────

API_URL = "https://restapi.jobs.krd/api/v1/public/jobs"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Referer": "https://jobs.krd/",
}

PARAMS = {
    "page_size": 40,
    "sort_by": "from_date",
    "sort_type": "desc",
    "type": "job",
}

TARGET = 220   # minimum rows to collect


# ── STEP 3: FETCH EACH PAGE ──────────────────────────────────

all_jobs = []
page = 1

print("Starting to scrape jobs.krd via API ...\n")

while len(all_jobs) < TARGET:

    PARAMS["page"] = page
    print(f"Fetching page {page} ...")

    response = requests.get(API_URL, headers=HEADERS, params=PARAMS, timeout=15)

    if response.status_code != 200:
        print(f"  Error {response.status_code}. Stopping.")
        break

    # ── STEP 4: EXTRACT DATA FROM EACH JOB ──────────────────

    data = response.json()
    jobs = data.get("data", [])

    if not jobs:
        print("  No more jobs found. Stopping.")
        break

    print(f"  Got {len(jobs)} jobs from page {page}.")

    for job in jobs:
        # Build salary string only when the employer chose to show it
        salary = "Not disclosed"
        if job.get("is_salary_visible") and job.get("min_pay"):
            currency_raw = job.get("currency") or {}
            currency = currency_raw.get("name", "") if isinstance(currency_raw, dict) else str(currency_raw)
            min_pay  = job.get("min_pay", "")
            max_pay  = job.get("max_pay", "")
            pay_type = (job.get("pay_type") or {}).get("name", "")
            salary = f"{min_pay}–{max_pay} {currency} / {pay_type}".strip()

        # Deadline: keep only the date part (YYYY-MM-DD)
        deadline = (job.get("to_date") or "")[:10] or "Unknown"

        # Individual job page URL
        unique_link = job.get("unique_link", "")
        url = f"https://jobs.krd/explore-jobs/{unique_link}" if unique_link else "Unknown"

        all_jobs.append({
            "title"   : job.get("job_title", "Unknown"),
            "company" : (job.get("company") or {}).get("legal_name", "Unknown"),
            "location": (job.get("location") or {}).get("name", "Unknown"),
            "category": (job.get("job_category") or {}).get("name", "Unknown"),
            "job_type": (job.get("job_type") or {}).get("name", "Unknown"),
            "deadline": deadline,
            "salary"  : salary,
            "url"     : url,
        })

    print(f"  Total collected so far: {len(all_jobs)}")

    # Stop if we've reached the last page
    if not data.get("next_page_url"):
        print("  Reached the last page. Stopping.")
        break

    page += 1
    time.sleep(1)   # be polite — add a delay between requests (from your slides!)

print(f"\nScraping complete. Total rows collected: {len(all_jobs)}")


# ── STEP 5: CLEAN AND SAVE TO CSV ───────────────────────────

df = pd.DataFrame(all_jobs)

# Remove duplicate rows (same URL = same job)
df.drop_duplicates(subset="url", inplace=True)

# Fill any empty cells with "Unknown"
df.fillna("Unknown", inplace=True)

# Save to CSV
df.to_csv("jobs_krd.csv", index=False, encoding="utf-8-sig")

print(f"Saved to jobs_krd.csv")
print(f"Rows   : {len(df)}")
print(f"Columns: {list(df.columns)}")
print("\nPreview of first 3 rows:")
print(df.head(3).to_string())


# ============================================================
#  DATA SOURCE NOTE
# ============================================================
#  Source  : https://jobs.krd/explore-jobs
#  API     : https://restapi.jobs.krd/api/v1/public/jobs
#  About   : A public job board for the Kurdistan Region of Iraq.
#
#  Columns:
#   title     - job position name
#   company   - name of the hiring company
#   location  - city or area of the job
#   category  - industry / job function category
#   job_type  - e.g. Full-time, Part-time, Remote
#   deadline  - application closing date (YYYY-MM-DD)
#   salary    - salary range if the employer disclosed it
#   url       - direct link to the job post
# ============================================================
