"""
Zoho Careers Scraper (v2)
-------------------------
Scrapes open jobs: title, location, qualification, link
Saves to: zoho_jobs/zoho_jobs.csv
Also saves zoho_jobs/debug_first_job.txt (page text of the first job) so you
can check what the page looks like if some fields are still empty.

Setup (one time):
    pip install playwright openpyxl
    playwright install chromium

Run:
    python zoho_job_scraper.py

Please check Zoho's robots.txt / Terms before scraping, and keep the delay.
"""

import csv
import os
import re
import time

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

CAREERS_URL = "https://careers.zohocorp.com/jobs/Careers"
OUTPUT_FOLDER = "zoho_jobs"
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, "zoho_jobs.csv")
EXCEL_FILE = os.path.join(OUTPUT_FOLDER, "zoho_jobs.xlsx")
DEBUG_FILE = os.path.join(OUTPUT_FOLDER, "debug_first_job.txt")
DELAY_SECONDS = 2

# job links look like /jobs/careers/2803000614929615/Software-Developers
JOB_LINK_RE = re.compile(r"/jobs/careers/\d+", re.I)  # numeric id only (skips /rss etc.)

QUAL_HEADINGS = [
    "qualification", "qualifications", "requirements", "requirement",
    "eligibility", "education", "who can apply", "skills required",
    "what we expect", "must have",
]

# case-sensitive on purpose, so normal words like "be" / "me" are not matched
DEGREE_RE = re.compile(
    r"\b(B\.?\s?E\b|B\.?\s?Tech\b|B\.?\s?TECH\b|M\.?\s?E\b|M\.?\s?Tech\b|M\.?\s?TECH\b|"
    r"BCA|MCA|B\.?\s?Sc\b|M\.?\s?Sc\b|B\.?\s?SC\b|M\.?\s?SC\b|BBA|MBA|"
    r"B\.?\s?Com\b|M\.?\s?Com\b|Diploma|Any Graduate|Any Degree|Bachelor|Master)",
)


def get_all_text(page):
    """Visible text from the main page and every iframe."""
    parts = []
    for frame in page.frames:
        try:
            parts.append(frame.inner_text("body"))
        except Exception:
            pass
    return "\n".join(parts)


def collect_job_links(page):
    page.goto(CAREERS_URL, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(4000)

    links = []

    def grab():
        for frame in page.frames:
            try:
                hrefs = frame.eval_on_selector_all(
                    "a[href]", "els => els.map(e => e.href)"
                )
            except Exception:
                continue
            for h in hrefs:
                h = h.split("?")[0]
                if JOB_LINK_RE.search(h) and h not in links:
                    links.append(h)

    try:
        page.wait_for_selector("a[href*='/jobs/' i]", timeout=15000)
    except Exception:
        pass
    grab()
    stagnant = 0
    for _ in range(40):
        before = len(links)
        page.mouse.wheel(0, 5000)
        page.wait_for_timeout(1200)

        # click load more / next page if present
        for frame in page.frames:
            try:
                btn = frame.query_selector("text=/^\\s*(load more|show more|next)\\s*$/i")
                if btn and btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(1500)
            except Exception:
                pass

        grab()
        stagnant = stagnant + 1 if len(links) == before else 0
        if stagnant >= 3:
            break

    # debug: dump every link on the listing page
    all_links = []
    for frame in page.frames:
        try:
            all_links += frame.eval_on_selector_all(
                "a[href]", "els => els.map(e => e.innerText.trim() + ' -> ' + e.href)")
        except Exception:
            pass
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    with open(os.path.join(OUTPUT_FOLDER, "debug_listing_links.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(all_links))
    return links


def parse_job_information(lines):
    """Zoho Recruit shows a 'Job Information' block: label line, then value line."""
    info = {}
    try:
        start = next(i for i, l in enumerate(lines) if l.lower() == "job information")
    except StopIteration:
        return info
    block = lines[start + 1:start + 40]
    i = 0
    while i < len(block) - 1:
        label = block[i].rstrip(":").strip()
        value = block[i + 1].strip()
        if len(label) < 30 and label.lower() in (
            "city", "state", "state/province", "province", "country",
            "region", "location", "work location", "industry",
            "job type", "experience", "salary",
        ):
            info[label.lower()] = value
            i += 2
        else:
            i += 1
    return info


def extract_location(lines, text):
    info = parse_job_information(lines)
    country = info.get("country") or info.get("region") or ""

    city = ""
    m = re.search(r"(?:work\s+)?location[s]?\s*[:\-]\s*([^\n]+)", text, re.I)
    if m:
        city = m.group(1).strip()
    elif info.get("city"):
        city = info["city"]
        if info.get("state") or info.get("state/province"):
            city += ", " + (info.get("state") or info.get("state/province"))

    if city and country and country.lower() not in city.lower():
        return f"{city}, {country}"
    return city or country or "Not mentioned"


def extract_qualification(lines, text):
    # 1) heading followed by lines
    for i, line in enumerate(lines):
        low = line.lower().rstrip(":").strip()
        if low in QUAL_HEADINGS:
            block = lines[i + 1:i + 8]
            return " | ".join(block)
        for k in QUAL_HEADINGS:
            if low.startswith(k) and ":" in line:
                val = line.split(":", 1)[1].strip()
                if val:
                    return val

    # 2) fallback: degrees mentioned anywhere in description
    found = []
    for m in DEGREE_RE.finditer(text):
        d = re.sub(r"\s+", "", m.group(0)).upper()
        if d not in found:
            found.append(d)
    if found:
        return ", ".join(found)
    return "Not mentioned"


def extract_description(lines):
    """Text under the 'Job Description' heading, until the next section starts."""
    stop_words = ("job information", "share job via email",
                  "share this job with your network", "view all jobs",
                  "visit website", "manage cookies")
    try:
        start = next(i for i, l in enumerate(lines) if l.lower() == "job description")
    except StopIteration:
        return "Not mentioned"

    desc = []
    for line in lines[start + 1:]:
        low = line.lower()
        if low in stop_words or ("interested" in low and len(line) < 25):
            break
        desc.append(line)
    return " ".join(desc) if desc else "Not mentioned"


def scrape_job(page, url, save_debug=False):
    page.goto(url, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(3000)

    text = get_all_text(page)
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    if save_debug:
        with open(DEBUG_FILE, "w", encoding="utf-8") as f:
            f.write(url + "\n\n" + text)

    title = ""
    for frame in page.frames:
        try:
            h = frame.query_selector("h1, h2")
            if h and h.inner_text().strip():
                title = h.inner_text().strip()
                break
        except Exception:
            pass
    if not title:
        title = page.title().strip()

    info = parse_job_information(lines)
    return {
        "title": title,
        "location": extract_location(lines, text),
        "qualification": extract_qualification(lines, text),
        "description": extract_description(lines),
        "industry": info.get("industry", "Not mentioned"),
        "job_type": info.get("job type", "Not mentioned"),
        "link": url,
    }


def save_excel(jobs, path):
    """Excel file where the description is shown as a wrapped paragraph."""
    columns = ["title", "location", "qualification", "description",
               "industry", "job_type", "link"]
    widths = [30, 22, 30, 90, 16, 12, 45]

    wb = Workbook()
    ws = wb.active
    ws.title = "Zoho Jobs"
    ws.append(columns)
    for job in jobs:
        ws.append([job[c] for c in columns])

    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = w

    wrap_top = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = wrap_top
    for cell in ws[1]:
        cell.font = Font(bold=True)

    try:
        wb.save(path)
    except PermissionError:
        path = os.path.join(OUTPUT_FOLDER, f"zoho_jobs_{int(time.time())}.xlsx")
        print("Excel file open-la irukku, so new file-la save panren...")
        wb.save(path)
    return path


def main():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    jobs = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        print("Collecting job links...")
        links = collect_job_links(page)
        print(f"Found {len(links)} jobs")

        for i, url in enumerate(links, 1):
            try:
                job = scrape_job(page, url, save_debug=(i == 1))
                jobs.append(job)
                print(f"[{i}/{len(links)}] {job['title']} | {job['location']} | {job['qualification'][:50]}")
            except PWTimeout:
                print(f"[{i}/{len(links)}] Timeout: {url}")
            except Exception as e:
                print(f"[{i}/{len(links)}] Error on {url}: {e}")
            time.sleep(DELAY_SECONDS)

        browser.close()

    out = OUTPUT_FILE
    try:
        f = open(out, "w", newline="", encoding="utf-8-sig")
    except PermissionError:
        out = os.path.join(OUTPUT_FOLDER, f"zoho_jobs_{int(time.time())}.csv")
        print("zoho_jobs.csv Excel-la open-la irukku, so new file-la save panren...")
        f = open(out, "w", newline="", encoding="utf-8-sig")
    with f:
        writer = csv.DictWriter(f, fieldnames=["title", "location", "qualification", "description", "industry", "job_type", "link"])
        writer.writeheader()
        writer.writerows(jobs)

    xlsx_path = save_excel(jobs, EXCEL_FILE)
    print(f"\nDone! Saved {len(jobs)} jobs to {out} and {xlsx_path}")


if __name__ == "__main__":
    main()