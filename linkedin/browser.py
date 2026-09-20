import sys
import time
from pathlib import Path
from urllib.parse import quote

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from config import (
    JOB_SEARCHES,
    MAX_JOB_AGE_DAYS,
    MIN_COMPANY_EMPLOYEES,
    MAX_SCROLLS,
    SCROLL_WAIT_SECONDS,
    CHROME_BINARY,
    CHROME_PROFILE,
    PAGE_LOAD_TIMEOUT_SECONDS,
    SCRIPT_TIMEOUT_SECONDS,
    SEARCH_NAVIGATION_RETRIES,
    EXCLUDED_COMPANIES,
)

from linkedin.jobs import (
    extract_job_cards,
    extract_job_details,
    extract_company_employee_count,
    is_excluded_company,
)

SEARCH_BASE_URL = "https://www.linkedin.com/jobs/search/"


class LinkedInSessionExpired(RuntimeError):
    """Raised when LinkedIn has redirected the browser to a login/authwall/checkpoint page."""


_AUTHWALL_URL_MARKERS = (
    "/login",
    "/uas/login",
    "/authwall",
    "/checkpoint",
    "/challenge",
)


def linkedin_session_valid(driver, log=True):
    """Return True only when the current page is not an obvious LinkedIn authwall/login.

    This is deliberately a session-state guard, not an anti-bot or verification bypass.
    """
    try:
        current_url = (driver.current_url or "").lower()
    except Exception:
        return False

    if any(marker in current_url for marker in _AUTHWALL_URL_MARKERS):
        if log:
            print(f"LINKEDIN SESSION BLOCKED: authentication/verification page detected: {driver.current_url}")
        return False

    try:
        # Strong login-page indicators only; avoid generic 'sign in' footer text.
        login_forms = driver.find_elements(
            By.CSS_SELECTOR,
            "form[action*='login'], input[name='session_key'], input[name='session_password']",
        )
        if login_forms:
            if log:
                print("LINKEDIN SESSION BLOCKED: login form detected on current page.")
            return False
    except Exception:
        pass

    return True


def require_linkedin_session(driver, context="LinkedIn page"):
    """Raise a distinct error instead of silently treating an authwall as zero results."""
    if not linkedin_session_valid(driver, log=True):
        raise LinkedInSessionExpired(
            f"LinkedIn session is unavailable while opening {context}. "
            """Please log in again in the configured Chrome profile and rerun."""
        )
    return True


def create_driver():
    options = Options()
    options.binary_location = CHROME_BINARY
    options.page_load_strategy = "eager"

    options.add_argument(f"--user-data-dir={CHROME_PROFILE}")
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    print("Starting Selenium Chrome...")

    driver = webdriver.Chrome(options=options)

    # This is the important fix for the 120-second hang.
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT_SECONDS)
    driver.set_script_timeout(SCRIPT_TIMEOUT_SECONDS)

    print("Chrome started!")
    print(f"Page-load timeout: {PAGE_LOAD_TIMEOUT_SECONDS}s")

    return driver


def safe_get(driver, url, label="page"):
    for attempt in range(1, SEARCH_NAVIGATION_RETRIES + 1):
        try:
            driver.get(url)
            require_linkedin_session(driver, label)
            return True

        except TimeoutException:
            print(
                f"Navigation timeout for {label} "
                f"(attempt {attempt}/{SEARCH_NAVIGATION_RETRIES}); "
                "continuing with whatever loaded."
            )
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass
            # A timeout is not necessarily fatal, but an authwall is.
            require_linkedin_session(driver, label)
            return False

        except WebDriverException as exc:
            print(
                f"Navigation error for {label} "
                f"(attempt {attempt}/{SEARCH_NAVIGATION_RETRIES}): {exc}"
            )
            if attempt < SEARCH_NAVIGATION_RETRIES:
                time.sleep(1)

    return False


def load_more_jobs(driver):
    print("\nLoading more LinkedIn jobs...")

    previous_count = 0
    stable_rounds = 0

    for scroll_number in range(1, MAX_SCROLLS + 1):
        try:
            cards = driver.find_elements(
                By.CSS_SELECTOR,
                "div.job-card-container",
            )
            current_count = len(cards)
        except Exception:
            current_count = 0

        print(f"Scroll {scroll_number}: {current_count} cards loaded")

        try:
            driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);"
            )
        except Exception:
            break

        time.sleep(SCROLL_WAIT_SECONDS)

        show_more_selectors = [
            "button.infinite-scroller__show-more-button",
            "button[aria-label*='Show more']",
            "button[aria-label*='more jobs']",
        ]

        clicked = False

        for selector in show_more_selectors:
            try:
                buttons = driver.find_elements(
                    By.CSS_SELECTOR,
                    selector,
                )

                for button in buttons:
                    if button.is_displayed() and button.is_enabled():
                        driver.execute_script(
                            "arguments[0].click();",
                            button,
                        )
                        print("Clicked Show More Jobs.")
                        time.sleep(SCROLL_WAIT_SECONDS)
                        clicked = True
                        break

                if clicked:
                    break

            except Exception:
                continue

        try:
            cards = driver.find_elements(
                By.CSS_SELECTOR,
                "div.job-card-container",
            )
            new_count = len(cards)
        except Exception:
            new_count = current_count

        print(f"After loading: {new_count} cards")

        if new_count == previous_count:
            stable_rounds += 1
        else:
            stable_rounds = 0

        previous_count = new_count

        if stable_rounds >= 2:
            print("No additional jobs loaded.")
            break

    return previous_count


def extract_all_loaded_jobs(driver):
    jobs = extract_job_cards(driver)

    unique_jobs = {}

    for job in jobs:
        job_id = job.get("job_id")
        if not job_id:
            continue

        company = job.get("company", "")
        if is_excluded_company(company):
            print(f"SKIP excluded company: {company}")
            continue

        unique_jobs[job_id] = job

    result = list(unique_jobs.values())

    print(f"\nUnique jobs extracted: {len(result)}")
    return result


def search_jobs(driver, keyword):
    print("\n========================================")
    print(f"SEARCH: {keyword}")
    print("========================================")

    search_url = (
        SEARCH_BASE_URL
        + f"?keywords={quote(keyword)}"
        + "&f_TPR=r604800"
    )

    loaded = safe_get(driver, search_url, f"search '{keyword}'")

    if not loaded:
        print(
            "Search navigation timed out/partially loaded. "
            "Checking page for available job cards."
        )

    time.sleep(3)

    require_linkedin_session(driver, f"job search '{keyword}'")

    try:
        print("URL:", driver.current_url)
        print("Title:", driver.title)
    except Exception:
        pass

    require_linkedin_session(driver, f"job search '{keyword}'")

    try:
        load_more_jobs(driver)
        jobs = extract_all_loaded_jobs(driver)
    except Exception as exc:
        print(f"Search extraction failed for '{keyword}': {exc}")
        return []

    print(f"Extracted {len(jobs)} jobs for search: {keyword}")

    recent_jobs = []

    for job in jobs:
        company_from_card = job.get("company", "")

        if is_excluded_company(company_from_card):
            print(f"STATUS: SKIP - excluded company: {company_from_card}")
            continue

        print("\n----------------------------------------")
        print(f"JOB: {job.get('title')}")
        print(f"Company: {company_from_card}")
        print(f"Job ID: {job.get('job_id')}")

        try:
            details = extract_job_details(driver, job)
            require_linkedin_session(driver, f"job details for {job.get('job_id')}")
        except Exception as exc:
            print(
                f"STATUS: SKIP - could not load job details "
                f"for {job.get('job_id')}: {exc}"
            )
            continue

        company = details.get("company", company_from_card)

        if is_excluded_company(company):
            print(f"STATUS: SKIP - excluded company: {company}")
            continue

        if (
            details.get("is_internship")
            or details.get("employment_type") == "Internship"
        ):
            print("STATUS: SKIP - internship")
            continue

        employment_type = details.get("employment_type", "Unknown")
        print("Employment type:", employment_type)

        if employment_type in [
            "Part-time",
            "Temporary",
            "Volunteer",
            "Contract",
        ]:
            print(f"STATUS: SKIP - {employment_type}")
            continue

        min_exp = details.get("min_experience_years")
        max_exp = details.get("max_experience_years")

        print("Experience:", details.get("experience"))

        if min_exp is not None and min_exp > 2:
            print("STATUS: SKIP - requires >2 years")
            continue

        if (
            min_exp is not None
            and max_exp is not None
            and max_exp < 1.5
        ):
            print("STATUS: SKIP - below experience range")
            continue

        age = details.get("job_age_days")
        print("Posting age:", age)

        if age is None:
            print("STATUS: SKIP - posting age unknown")
            continue

        if age > MAX_JOB_AGE_DAYS:
            print(f"STATUS: SKIP - posted {age} days ago")
            continue

        print("Salary:", details.get("salary_text", ""))

        company_url = details.get("company_url", "")

        if company_url:
            try:
                company_info = extract_company_employee_count(
                    driver,
                    company_url,
                )
                require_linkedin_session(driver, f"company details for {company}")
                details.update(company_info)
            except Exception as exc:
                print(f"Company size lookup failed: {exc}")

        min_employees = details.get("min_employees")
        max_employees = details.get("max_employees")

        print("Company employees:", details.get("employee_text", ""))

        if min_employees is None:
            print("STATUS: KEEP - company size unknown")
        elif min_employees >= MIN_COMPANY_EMPLOYEES:
            print("Company size filter: PASS")
        elif (
            max_employees is not None
            and max_employees >= MIN_COMPANY_EMPLOYEES
        ):
            print("Company size: borderline range; keeping")
        else:
            print("STATUS: SKIP - company <200 employees")
            continue

        details["search_term"] = keyword
        recent_jobs.append(details)

        print("STATUS: FINAL KEEP")

    print("\n----------------------------------------")
    print(f"SEARCH COMPLETE: {keyword}")
    print(f"Final jobs: {len(recent_jobs)}")

    return recent_jobs


def search_all_jobs(driver):
    all_jobs = {}

    print("\n========================================")
    print("RUNNING ALL JOB SEARCHES")
    print(f"Total searches: {len(JOB_SEARCHES)}")
    print("========================================")

    for index, keyword in enumerate(JOB_SEARCHES, start=1):
        print(f"\n[{index}/{len(JOB_SEARCHES)}]")

        try:
            jobs = search_jobs(driver, keyword)

            for job in jobs:
                job_id = job.get("job_id")
                if not job_id:
                    continue

                company = job.get("company", "")
                if is_excluded_company(company):
                    continue

                if job_id not in all_jobs:
                    job["search_terms"] = keyword
                    all_jobs[job_id] = job
                else:
                    existing = all_jobs[job_id]
                    existing_terms = str(
                        existing.get("search_terms", "")
                    )

                    terms = {
                        value.strip()
                        for value in existing_terms.split(",")
                        if value.strip()
                    }
                    terms.add(keyword)
                    existing["search_terms"] = ", ".join(sorted(terms))

        except LinkedInSessionExpired:
            print(
                f"STOPPING ALL SEARCHES: LinkedIn session unavailable during '{keyword}'."
            )
            raise
        except Exception as error:
            print(
                f"Search failed for '{keyword}': {error}"
            )
            continue

    result = list(all_jobs.values())

    print("\n========================================")
    print("ALL SEARCHES COMPLETE")
    print(f"Unique final jobs: {len(result)}")
    print("========================================")

    return result


def main():
    driver = create_driver()

    try:
        jobs = search_all_jobs(driver)

        print("\nFINAL JOB LIST")

        for index, job in enumerate(jobs, start=1):
            print(f"\n{index}. {job.get('title')}")
            print("   Company:", job.get("company"))
            print("   Job ID:", job.get("job_id"))
            print("   Age:", job.get("job_age_days"))
            print("   Search:", job.get("search_terms"))
            print("   URL:", job.get("job_url"))

        input("\nPress ENTER to close Chrome...")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
