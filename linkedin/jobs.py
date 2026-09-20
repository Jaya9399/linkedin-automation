import re
import time

from selenium.common.exceptions import TimeoutException

from selenium.webdriver.common.by import By

from config import EXCLUDED_COMPANIES, DETAIL_WAIT_SECONDS, COMPANY_WAIT_SECONDS



# ============================================================
# SAFE NAVIGATION
# ============================================================

def is_excluded_company(company):
    name = clean(company).lower()
    return any(
        clean(excluded).lower() in name
        for excluded in EXCLUDED_COMPANIES
        if clean(excluded)
    )


def safe_driver_get(driver, url, timeout_message=""):
    """
    Navigate without allowing a slow LinkedIn response to block the whole
    search run for Selenium's default 120-second command timeout.
    A navigation timeout is treated as a partial-load condition; the caller
    can continue extracting whatever LinkedIn rendered.
    """
    try:
        driver.get(url)
        return True
    except TimeoutException:
        label = timeout_message or url
        print(f"Navigation timeout; continuing with loaded page: {label}")
        try:
            driver.execute_script("window.stop();")
        except Exception:
            pass
        return False
    except Exception as exc:
        print(f"Navigation failed: {exc}")
        return False

# ============================================================
# CLEAN
# ============================================================

def clean(value):

    return re.sub(
        r"\s+",
        " ",
        value or "",
    ).strip()


# ============================================================
# JOB ID
# ============================================================

def extract_job_id(url):

    if not url:
        return ""

    match = re.search(
        r"/jobs/view/(\d+)",
        url,
    )

    if match:
        return match.group(1)

    return ""


# ============================================================
# POSTED TEXT
# ============================================================

def extract_posted_text(card):

    selectors = [
        "time",
        "span.job-card-container__footer-item",
        "li.job-card-container__footer-item",
    ]

    for selector in selectors:

        try:

            elements = card.find_elements(
                By.CSS_SELECTOR,
                selector,
            )

            for element in elements:

                text = clean(
                    element.text
                )

                if not text:
                    continue

                low = text.lower()

                if any(
                    word in low
                    for word in [
                        "ago",
                        "just now",
                        "today",
                        "yesterday",
                    ]
                ):

                    return text

        except Exception:
            continue

    return ""


# ============================================================
# AGE
# ============================================================

def calculate_job_age(
    posted_text
):

    if not posted_text:
        return None

    text = posted_text.lower()

    if (
        "just now" in text
        or "today" in text
    ):

        return 0

    if "yesterday" in text:

        return 1

    hours = re.search(
        r"(\d+)\s*(hour|hours|hr|hrs)",
        text,
    )

    if hours:
        return 0

    days = re.search(
        r"(\d+)\s*(day|days)",
        text,
    )

    if days:
        return int(
            days.group(1)
        )

    weeks = re.search(
        r"(\d+)\s*(week|weeks)",
        text,
    )

    if weeks:

        return int(
            weeks.group(1)
        ) * 7

    months = re.search(
        r"(\d+)\s*(month|months)",
        text,
    )

    if months:

        return int(
            months.group(1)
        ) * 30

    return None


# ============================================================
# INTERNSHIP
# ============================================================

def is_internship(job):

    title = (
        job.get(
            "title",
            "",
        )
        .lower()
    )

    description = (
        job.get(
            "description",
            "",
        )
        .lower()
    )

    employment_type = (
        job.get(
            "employment_type",
            "",
        )
        .lower()
    )

    if employment_type == "internship":
        return True

    internship_keywords = [
        "intern",
        "internship",
        "trainee",
        "apprentice",
        "graduate intern",
    ]

    for keyword in internship_keywords:

        if keyword in title:
            return True

    internship_phrases = [
        "this is an internship",
        "intern position",
        "internship position",
        "intern role",
        "as an intern",
        "summer internship",
    ]

    for phrase in internship_phrases:

        if phrase in description:
            return True

    return False


# ============================================================
# EMPLOYMENT TYPE
# ============================================================

def extract_employment_type(
    description
):

    text = (
        description or ""
    ).lower()

    if (
        "full-time" in text
        or "full time" in text
    ):

        return "Full-time"

    if (
        "part-time" in text
        or "part time" in text
    ):

        return "Part-time"

    if "internship" in text:
        return "Internship"

    if "intern" in text:
        return "Internship"

    if "temporary" in text:
        return "Temporary"

    if "volunteer" in text:
        return "Volunteer"

    if "contract" in text:
        return "Contract"

    return "Unknown"


# ============================================================
# SALARY
# ============================================================

def extract_salary(
    description
):

    if not description:

        return {
            "salary_text": "",
            "min_lpa": None,
            "max_lpa": None,
        }

    text = description.lower()

    # 8-12 LPA

    pattern = re.compile(
        r"₹?\s*(\d+(?:\.\d+)?)"
        r"\s*(?:-|to)\s*"
        r"₹?\s*(\d+(?:\.\d+)?)"
        r"\s*lpa"
    )

    match = pattern.search(
        text
    )

    if match:

        return {
            "salary_text": match.group(0),
            "min_lpa": float(
                match.group(1)
            ),
            "max_lpa": float(
                match.group(2)
            ),
        }

    # Single LPA

    pattern = re.compile(
        r"₹?\s*(\d+(?:\.\d+)?)"
        r"\s*lpa"
    )

    match = pattern.search(
        text
    )

    if match:

        value = float(
            match.group(1)
        )

        return {
            "salary_text": match.group(0),
            "min_lpa": value,
            "max_lpa": value,
        }

    # Indian salary range

    pattern = re.compile(
        r"₹\s*([\d,]+)"
        r"\s*(?:-|to)\s*"
        r"₹?\s*([\d,]+)"
    )

    match = pattern.search(
        text
    )

    if match:

        minimum = int(
            match.group(1).replace(
                ",",
                "",
            )
        )

        maximum = int(
            match.group(2).replace(
                ",",
                "",
            )
        )

        return {
            "salary_text": match.group(0),
            "min_lpa": minimum / 100000,
            "max_lpa": maximum / 100000,
        }

    return {
        "salary_text": "",
        "min_lpa": None,
        "max_lpa": None,
    }


# ============================================================
# EXPERIENCE
# ============================================================

def extract_experience(
    description
):

    if not description:

        return {
            "experience_text": "",
            "min_experience_years": None,
            "max_experience_years": None,
        }

    text = description.lower()

    # Range

    patterns = [
        (
            r"(\d+(?:\.\d+)?)\s*[-–]\s*"
            r"(\d+(?:\.\d+)?)\s*"
            r"(?:years?|yrs?)"
        ),

        (
            r"(\d+(?:\.\d+)?)\s+to\s+"
            r"(\d+(?:\.\d+)?)\s*"
            r"(?:years?|yrs?)"
        ),
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if match:

            return {
                "experience_text": match.group(0),
                "min_experience_years": float(
                    match.group(1)
                ),
                "max_experience_years": float(
                    match.group(2)
                ),
            }

    # Minimum

    patterns = [
        r"(\d+(?:\.\d+)?)\s*\+\s*(?:years?|yrs?)",
        r"minimum\s+of\s+(\d+(?:\.\d+)?)\s*(?:years?|yrs?)",
        r"minimum\s+(\d+(?:\.\d+)?)\s*(?:years?|yrs?)",
        r"at\s+least\s+(\d+(?:\.\d+)?)\s*(?:years?|yrs?)",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if match:

            return {
                "experience_text": match.group(0),
                "min_experience_years": float(
                    match.group(1)
                ),
                "max_experience_years": None,
            }

    # Exact

    patterns = [
        r"(\d+(?:\.\d+)?)\s*(?:years?|yrs?)\s+of\s+experience",
        r"(\d+(?:\.\d+)?)\s*(?:years?|yrs?)\s+experience",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if match:

            value = float(
                match.group(1)
            )

            return {
                "experience_text": match.group(0),
                "min_experience_years": value,
                "max_experience_years": value,
            }

    return {
        "experience_text": "",
        "min_experience_years": None,
        "max_experience_years": None,
    }


# ============================================================
# COMPANY URL
# ============================================================

def extract_company_url_from_card(
    card
):

    selectors = [
        "a[href*='/company/']",
        "a.job-card-container__company-name",
        "a.artdeco-entity-lockup__subtitle",
    ]

    for selector in selectors:

        try:

            elements = card.find_elements(
                By.CSS_SELECTOR,
                selector,
            )

            for element in elements:

                href = (
                    element.get_attribute(
                        "href"
                    )
                    or ""
                )

                if "/company/" in href:

                    return href.split(
                        "?"
                    )[0]

        except Exception:
            continue

    return ""


# ============================================================
# JOB CARDS
# ============================================================

def extract_job_cards(
    driver
):

    jobs = []

    time.sleep(2)

    cards = driver.find_elements(
        By.CSS_SELECTOR,
        "div.job-card-container",
    )

    print(
        f"Found {len(cards)} job cards"
    )

    for card in cards:

        try:

            title_element = card.find_element(
                By.CSS_SELECTOR,
                "a.job-card-list__title--link",
            )

            title = clean(
                title_element.text
            )

            title = re.sub(
                r"\s+with verification$",
                "",
                title,
                flags=re.IGNORECASE,
            ).strip()

            job_url = (
                title_element.get_attribute(
                    "href"
                )
                or ""
            )

            job_id = extract_job_id(
                job_url
            )

            if not job_id:
                continue

            try:

                company = clean(
                    card.find_element(
                        By.CSS_SELECTOR,
                        "div.artdeco-entity-lockup__subtitle",
                    ).text
                )

            except Exception:

                company = ""

            if is_excluded_company(company):
                print(f"STATUS: SKIP - excluded company: {company}")
                continue

            company_url = (
                extract_company_url_from_card(
                    card
                )
            )

            try:

                location = clean(
                    card.find_element(
                        By.CSS_SELECTOR,
                        "div.artdeco-entity-lockup__caption",
                    ).text
                )

            except Exception:

                location = ""

            posted_text = (
                extract_posted_text(
                    card
                )
            )

            job_age_days = (
                calculate_job_age(
                    posted_text
                )
            )

            jobs.append(
                {
                    "job_id": job_id,
                    "title": title,
                    "company": company,
                    "company_url": company_url,
                    "location": location,
                    "posted_text": posted_text,
                    "job_age_days": job_age_days,
                    "job_url": (
                        f"https://www.linkedin.com/"
                        f"jobs/view/{job_id}/"
                    ),
                }
            )

        except Exception:

            continue

    return jobs


# ============================================================
# COMPANY EMPLOYEE COUNT
# ============================================================

def extract_company_employee_count(
    driver,
    company_url,
):

    result = {
        "employee_text": "",
        "min_employees": None,
        "max_employees": None,
    }

    if not company_url:
        return result

    try:

        print(
            f"Opening company page: "
            f"{company_url}"
        )

        safe_driver_get(driver, company_url, 'company page')

        time.sleep(COMPANY_WAIT_SECONDS)

        page_text = clean(
            driver.find_element(
                By.TAG_NAME,
                "body",
            ).text
        )

        patterns = [
            r"(\d[\d,]*)\s*[-–]\s*(\d[\d,]*)\s+employees",
            r"(\d[\d,]*)\s+to\s+(\d[\d,]*)\s+employees",
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                page_text,
                re.IGNORECASE,
            )

            if match:

                minimum = int(
                    match.group(1).replace(
                        ",",
                        "",
                    )
                )

                maximum = int(
                    match.group(2).replace(
                        ",",
                        "",
                    )
                )

                result[
                    "employee_text"
                ] = match.group(0)

                result[
                    "min_employees"
                ] = minimum

                result[
                    "max_employees"
                ] = maximum

                return result

        plus = re.search(
            r"(\d[\d,]*)\+\s+employees",
            page_text,
            re.IGNORECASE,
        )

        if plus:

            result[
                "employee_text"
            ] = plus.group(0)

            result[
                "min_employees"
            ] = int(
                plus.group(1).replace(
                    ",",
                    "",
                )
            )

            return result

    except Exception as exc:

        print(
            "Company size extraction failed:",
            exc,
        )

    return result


# ============================================================
# JOB DETAILS
# ============================================================

def extract_job_details(
    driver,
    job,
):

    print(
        f"\nOpening job: "
        f"{job['job_id']}"
    )

    safe_driver_get(driver, job["job_url"], f'job {job["job_id"]}')

    time.sleep(DETAIL_WAIT_SECONDS)

    details = job.copy()

    try:

        page_text = clean(
            driver.find_element(
                By.TAG_NAME,
                "body",
            ).text
        )

    except Exception:

        page_text = ""

    description = ""

    selectors = [
        "div.jobs-description__content",
        "div.jobs-box__html-content",
        "div.jobs-description-content__text",
        "article.jobs-description__container",
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector,
            )

            for element in elements:

                text = clean(
                    element.text
                )

                if len(text) > len(
                    description
                ):

                    description = text

        except Exception:

            continue

    if not description:

        description = page_text

    details[
        "description"
    ] = description

    salary = extract_salary(
        description
    )

    details[
        "salary_text"
    ] = salary[
        "salary_text"
    ]

    details[
        "min_lpa"
    ] = salary[
        "min_lpa"
    ]

    details[
        "max_lpa"
    ] = salary[
        "max_lpa"
    ]

    details[
        "employment_type"
    ] = extract_employment_type(
        description
    )

    details[
        "is_internship"
    ] = is_internship(
        details
    )

    # --------------------------------------------------------
    # Posted text
    # --------------------------------------------------------

    posted_text = ""

    selectors = [
        "span.posted-time-ago__text",
        "span.jobs-unified-top-card__posted-date",
        "div.jobs-unified-top-card__primary-description-container",
        "div.t-black--light",
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector,
            )

            for element in elements:

                text = clean(
                    element.text
                )

                low = text.lower()

                if any(
                    word in low
                    for word in [
                        "ago",
                        "just now",
                        "today",
                        "yesterday",
                        "week",
                        "month",
                        "day",
                    ]
                ):

                    posted_text = text

                    break

            if posted_text:
                break

        except Exception:

            continue

    if not posted_text:

        patterns = [
            r"Posted\s+(\d+\s+\w+\s+ago)",
            r"(\d+\s+\w+\s+ago)",
            r"(Just now)",
            r"(Today)",
            r"(Yesterday)",
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                page_text,
                re.IGNORECASE,
            )

            if match:

                posted_text = match.group(1)

                break

    details[
        "posted_text"
    ] = posted_text

    details[
        "job_age_days"
    ] = calculate_job_age(
        posted_text
    )

    # --------------------------------------------------------
    # Experience
    # --------------------------------------------------------

    experience = extract_experience(
        description
    )

    details[
        "experience"
    ] = experience[
        "experience_text"
    ]

    details[
        "min_experience_years"
    ] = experience[
        "min_experience_years"
    ]

    details[
        "max_experience_years"
    ] = experience[
        "max_experience_years"
    ]

    # --------------------------------------------------------
    # Company
    # --------------------------------------------------------

    try:

        company_links = driver.find_elements(
            By.CSS_SELECTOR,
            "a[href*='/company/']",
        )

        for link in company_links:

            href = (
                link.get_attribute(
                    "href"
                )
                or ""
            )

            if "/company/" in href:

                details[
                    "company_url"
                ] = href.split("?")[0]

                text = clean(
                    link.text
                )

                if text:

                    details[
                        "company"
                    ] = text

                break

    except Exception:
        pass

    try:

        company_element = driver.find_element(
            By.CSS_SELECTOR,
            "a.topcard__org-name-link",
        )

        company_name = clean(
            company_element.text
        )

        company_url = (
            company_element.get_attribute(
                "href"
            )
            or ""
        )

        if company_name:

            details[
                "company"
            ] = company_name

        if company_url:

            details[
                "company_url"
            ] = company_url.split(
                "?"
            )[0]

    except Exception:
        pass

    return details