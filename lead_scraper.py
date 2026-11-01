import csv
import time
import re
from urllib.parse import quote_plus
from playwright.sync_api import sync_playwright

SEARCH_QUERIES = [
    "Data Analytics Intern",
    "AI Intern",
    "Machine Learning Intern",
    "Data Science Intern",
    "Python Intern"
]
LOCATIONS = ["Selangor", "Kuala Lumpur"]
OUTPUT_CSV = "scraped_leads.csv"

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

def scrape_jobstreet():
    results = []
    seen_urls = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        for query in SEARCH_QUERIES:
            for loc in LOCATIONS:
                slug_q = query.lower().replace(" ", "-")
                slug_l = loc.lower().replace(" ", "-")
                url = f"https://my.jobstreet.com/{slug_q}-jobs/in-{slug_l}"
                print(f"[*] Scanning: {url}")
                try:
                    page.goto(url, timeout=30000)
                    page.wait_for_timeout(3000)

                    cards = page.locator("article[data-automation='normalJob']").all()
                    for card in cards[:10]:
                        try:
                            title_el = card.locator("[data-automation='jobTitle']").first
                            comp_el = card.locator("[data-automation='jobCompany']").first
                            loc_el = card.locator("[data-automation='jobLocation']").first

                            title = title_el.inner_text().strip() if title_el.count() else ""
                            company = comp_el.inner_text().strip() if comp_el.count() else ""
                            location = loc_el.inner_text().strip() if loc_el.count() else loc
                            href = title_el.get_attribute("href") if title_el.count() else ""
                            full_url = f"https://my.jobstreet.com{href}" if href and href.startswith("/") else href

                            if not title or full_url in seen_urls:
                                continue
                            seen_urls.add(full_url)

                            # Visit job detail page to check for direct HR email in description
                            hr_email = ""
                            if full_url:
                                detail_page = context.new_page()
                                try:
                                    detail_page.goto(full_url, timeout=20000)
                                    detail_page.wait_for_timeout(1500)
                                    body_text = detail_page.locator("body").inner_text()
                                    found_emails = EMAIL_REGEX.findall(body_text)
                                    valid_emails = [
                                        e for e in found_emails 
                                        if not any(x in e.lower() for x in ["jobstreet", "sentry", "example", "noreply"])
                                    ]
                                    if valid_emails:
                                        hr_email = valid_emails[0]
                                except Exception:
                                    pass
                                finally:
                                    detail_page.close()

                            results.append({
                                "Company_Name": company,
                                "Target_Role": title,
                                "Location": location,
                                "Lang_Pref": "EN",
                                "HR_Name": "Hiring Team",
                                "HR_Role": "Talent Acquisition",
                                "Company_Address": location,
                                "HR_Email": hr_email,
                                "Channel": full_url,
                                "Status": "Ready" if hr_email else "Pending Email",
                                "Notes": f"Scraped for query: {query}"
                            })
                            print(f"    + Found: {company} | {title} | Email: {hr_email or 'None (Portal)'}")
                        except Exception as e:
                            continue
                except Exception as e:
                    print(f"[!] Error loading {url}: {e}")

        browser.close()

    if results:
        fieldnames = [
            "Company_Name", "Target_Role", "Location", "Lang_Pref",
            "HR_Name", "HR_Role", "Company_Address", "HR_Email",
            "Channel", "Status", "Notes"
        ]
        with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
        print(f"\n[SUCCESS] Saved {len(results)} internship listings to {OUTPUT_CSV}")

if __name__ == "__main__":
    scrape_jobstreet()
