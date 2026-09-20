import os
import re
import time
import json
from urllib.parse import quote

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    WebDriverException,
    StaleElementReferenceException,
)

from config import (
    PEOPLE_ROLES,
    HIGH_VALUE_ROLES,
    VERY_HIGH_VALUE_ROLES,
    MAX_PEOPLE_PER_JOB,
    MAX_INITIAL_REFERRAL_CONTACTS,
    MAX_CONNECTION_REQUESTS_PER_RUN,
    MAX_REFERRAL_MESSAGES_PER_RUN,
    ACTIVE_MAX_DAYS,
    MIN_RELEVANCE_SCORE,
    PEOPLE_WAIT_SECONDS,
)
from storage.excel import (
    upsert_contact,
    get_contacts_for_job,
    referral_already_received,
    add_message,
    add_activity,
    mark_message_sent,
    mark_referral_requested,
    mark_connection_accepted,
    _read_sheet,
)
from messages.templates import (
    referral_message,
    connection_message,
    prepare_referral_already_received_message,
)


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _safe_storage_call(label, func, *args, **kwargs):
    """Call storage safely so a successful LinkedIn action is not reported as failed."""
    try:
        return func(*args, **kwargs)
    except Exception as exc:
        print(f"Storage warning ({label}): {exc}")
        return None

# ---------------------------------------------------------------------------
# Persistent referral-send safety
# ---------------------------------------------------------------------------
# "started" is intentionally fail-closed: once the automation has entered the
# referral composer and is about to type, that Job ID + Person is permanently
# locked from automatic retry. This protects against Ctrl+C/crashes during
# typing, before Excel can be updated.
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
REFERRAL_LOCK_FILE = os.path.join(PROJECT_ROOT, "storage", "referral_send_locks.json")


def _referral_lock_key(job_id, person):
    return f"{clean(job_id)}::{clean(person).casefold()}"


def _load_referral_send_locks():
    try:
        if not os.path.isfile(REFERRAL_LOCK_FILE):
            return {}
        with open(REFERRAL_LOCK_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception as exc:
        print(f"Referral safety warning: could not read lock file: {exc}")
        return None


def _save_referral_send_locks(data):
    directory = os.path.dirname(REFERRAL_LOCK_FILE)
    if directory:
        os.makedirs(directory, exist_ok=True)

    tmp = REFERRAL_LOCK_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, REFERRAL_LOCK_FILE)


def _get_referral_send_lock(job_id, person):
    data = _load_referral_send_locks()
    if data is None:
        return "__SAFETY_STATE_UNREADABLE__"
    return data.get(_referral_lock_key(job_id, person))


def _claim_referral_send_lock(job_id, person):
    """Persist a pre-send lock before referral text is typed."""
    data = _load_referral_send_locks()
    if data is None:
        print(
            f"Referral BLOCKED for {person}: persistent safety state could "
            "not be read. No message will be typed."
        )
        return False

    key = _referral_lock_key(job_id, person)
    existing = data.get(key)
    if existing:
        print(
            f"Skipping referral for {person}: persistent send lock already "
            f"exists for Job ID {job_id} "
            f"(state={existing.get('state', 'unknown')})."
        )
        return False

    data[key] = {
        "state": "started",
        "job_id": clean(job_id),
        "person": clean(person),
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }

    try:
        _save_referral_send_locks(data)
        print(
            f"Referral send lock claimed for {person} | job={job_id}. "
            "No later run will auto-retry this referral."
        )
        return True
    except Exception as exc:
        print(
            f"Referral BLOCKED for {person}: could not persist send lock: {exc}"
        )
        return False


def _mark_referral_send_lock(job_id, person, state, **extra):
    """Update final state without ever deleting the persistent lock."""
    data = _load_referral_send_locks()
    if data is None:
        print(
            f"Referral safety warning: could not update lock for "
            f"{person} | job={job_id}."
        )
        return False

    key = _referral_lock_key(job_id, person)
    entry = data.get(key, {})
    if not isinstance(entry, dict):
        entry = {}

    entry.update({
        "state": state,
        "job_id": clean(job_id),
        "person": clean(person),
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    })
    entry.update(extra)
    data[key] = entry

    try:
        _save_referral_send_locks(data)
        return True
    except Exception as exc:
        print(
            f"Referral safety warning: could not persist state={state} "
            f"for {person} | job={job_id}: {exc}"
        )
        return False


def role_priority(role):
    role = clean(role).lower()

    if "ceo" in role or "chief executive" in role:
        return 110
    if "founder" in role or "co-founder" in role:
        return 108
    if "vp " in role or role.startswith("vp") or "vice president" in role:
        return 105
    if "director" in role or "head of" in role:
        return 103
    if "hiring manager" in role:
        return 100
    if "engineering manager" in role:
        return 95
    if "recruiter" in role or "talent acquisition" in role or "talent partner" in role:
        return 90
    if "principal software engineer" in role:
        return 85
    if "staff software engineer" in role:
        return 82
    if "lead software engineer" in role or "tech lead" in role:
        return 78
    if "senior software engineer" in role:
        return 75
    if "software engineer" in role:
        return 55
    if "developer" in role or "engineer" in role:
        return 45

    return 20


def is_high_value_role(role):
    return role_priority(role) >= 75


def connection_priority(degree):
    degree = clean(degree).lower()

    if "1st" in degree or "first" in degree:
        return 30
    if "2nd" in degree or "second" in degree:
        return 20
    if "3rd" in degree or "third" in degree:
        return 10

    return 5


def activity_age_days(activity):
    """
    Convert visible LinkedIn activity text into an approximate age.

    Examples:
        "Posted 2d ago" -> 2
        "Posted 2w ago" -> 14
        "Posted 1mo ago" -> 30
        "Recently posted" -> 7

    None means LinkedIn did not expose enough evidence to call the person active.
    """
    text = clean(activity).lower()

    if not text:
        return None

    if any(x in text for x in ("just now", "today", "active today", "recently posted", "posted recently")):
        return 0

    patterns = [
        (r"(\d+)\s*(?:m|min|mins|minute|minutes)\b", 0),
        (r"(\d+)\s*(?:h|hr|hrs|hour|hours)\b", 0),
        (r"(\d+)\s*(?:d|day|days)\b", 1),
        (r"(\d+)\s*(?:w|wk|wks|week|weeks)\b", 7),
        (r"(\d+)\s*(?:mo|mos|month|months)\b", 30),
        (r"(\d+)\s*(?:y|yr|yrs|year|years)\b", 365),
    ]

    for pattern, multiplier in patterns:
        match = re.search(pattern, text)
        if match:
            value = int(match.group(1))
            return value * multiplier

    # "posted" alone is not enough evidence. Do not contact such a person.
    return None

def is_active_contact(contact):
    """
    Return True when activity is recent enough OR activity is unknown.

    Unknown activity should not eliminate a potentially valuable employee.
    It is treated as lower priority rather than automatically rejected.
    """

    activity = clean(
        contact.get(
            "activity",
            contact.get("Activity", ""),
        )
    )

    if not activity:
        return True

    age = activity_age_days(activity)

    if age is None:
        return True

    return age <= ACTIVE_MAX_DAYS

def activity_priority(activity):
    age = activity_age_days(activity)

    if age is None:
        return 0

    if age <= 3:
        return 30
    if age <= 7:
        return 25
    if age <= 14:
        return 20
    if age <= ACTIVE_MAX_DAYS:
        return 10

    return 0


def contact_score(contact):
    return (
        role_priority(contact.get("role", contact.get("Role", "")))
        + connection_priority(
            contact.get(
                "connection_degree",
                contact.get("Connection Degree", ""),
            )
        )
        + activity_priority(
            contact.get("activity", contact.get("Activity", ""))
        )
    )


def referral_priority_score(contact):
    return contact_score(contact)


def _url(contact):
    return clean(
        contact.get("linkedin_url", contact.get("LinkedIn URL", ""))
    ).split("?")[0].rstrip("/")


def _person(contact):
    return clean(contact.get("person", contact.get("Person", "")))


def _extract_card(card):
    """Extract a person from a LinkedIn People result card.

    LinkedIn frequently puts the person's name in aria-label. The previous
    implementation only parsed role/degree/activity when aria-label was empty,
    which silently discarded most valid cards. Parse all card metadata
    independently of how the name was obtained.
    """
    try:
        links = card.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
    except (WebDriverException, StaleElementReferenceException):
        return None

    if not links:
        return None

    try:
        href = clean(links[0].get_attribute("href"))
    except WebDriverException:
        return None

    url = _url({"linkedin_url": href})
    if not url:
        return None

    lines = []
    try:
        lines = [
            clean(x)
            for x in card.text.splitlines()
            if clean(x)
        ]
    except WebDriverException:
        pass

    person = ""
    try:
        aria = clean(links[0].get_attribute("aria-label"))
        if aria:
            # Some aria labels contain "Name • 2nd"; keep only the name.
            person = re.split(
                r"\s+[•·]\s+|\s+\|\s+|\s+2nd\s+|\s+3rd\s+|\s+1st\s+",
                aria,
                maxsplit=1,
                flags=re.I,
            )[0].strip()
    except WebDriverException:
        pass

    if not person:
        try:
            raw_person = clean(links[0].text)
        except WebDriverException:
            raw_person = ""

        person = re.split(
            r"\s+[•·]\s+|\s+\|\s+|\s+2nd\s+|\s+3rd\s+|\s+1st\s+",
            raw_person,
            maxsplit=1,
            flags=re.I,
        )[0].strip()

    if not person and lines:
        person = lines[0]

    role = ""

    # Parse the headline/designation from the card independently of the name.
    # Keep the configured referral roles as the primary signal, but recognize
    # common variants such as "Software Engineer, ..." and "Talent Partner".
    fragments = [
        "ceo",
        "chief executive",
        "founder",
        "co-founder",
        "vp ",
        "vice president",
        "director",
        "head of",
        "hiring manager",
        "engineering manager",
        "engineering lead",
        "technical recruiter",
        "recruiter",
        "recruiting",
        "talent acquisition",
        "talent partner",
        "talent management",
        "principal software engineer",
        "principal engineer",
        "staff software engineer",
        "staff engineer",
        "lead software engineer",
        "lead engineer",
        "tech lead",
        "senior software engineer",
        "software engineer",
        "software developer",
        "developer",
        "engineer",
    ]

    for line in lines:
        low = line.casefold()

        # Do not treat the person's name itself as their role.
        if person and low == person.casefold():
            continue

        if any(fragment in low for fragment in fragments):
            role = line
            break

    degree = ""
    for line in lines:
        low = line.casefold()

        if re.search(r"\b1st\b|\bfirst\b", low):
            degree = "1st"
            break
        if re.search(r"\b2nd\b|\bsecond\b", low):
            degree = "2nd"
            break
        if re.search(r"\b3rd\b|\bthird\b", low):
            degree = "3rd"
            break

    activity = next(
        (
            x
            for x in lines
            if any(
                k in x.casefold()
                for k in (
                    "active",
                    "posted",
                    "recent",
                    "engaged",
                )
            )
        ),
        "",
    )

    count = ""
    try:
        match = re.search(
            r"(\d[\d,]*)\+?\s+connections?",
            card.text,
            re.I,
        )
        if match:
            count = int(match.group(1).replace(",", ""))
    except WebDriverException:
        pass

    if not person:
        return None

    return {
        "person": person,
        "linkedin_url": url,
        "role": role,
        "connection_degree": degree,
        "connection_count": count,
        "activity": activity,
    }


def _collect_people_cards(driver):
    """
    Collect LinkedIn people result cards using several DOM patterns.

    LinkedIn changes its result-card HTML frequently, so we first look
    for known result containers and then fall back to profile links and
    their nearest useful ancestor.

    We intentionally collect people first and apply role/activity filters
    later. This prevents valid employees from being discarded too early.
    """

    found = []
    seen_urls = set()

    selectors = (
        "li.reusable-search__result-container",
        "div.entity-result",
        "div.entity-result__item",
        "li[data-chameleon-result-urn]",
        "div[data-view-name='people-search-result']",
        "div[data-view-name*='search-result']",
    )

    def add_card(card):
        try:
            links = card.find_elements(
                By.CSS_SELECTOR,
                "a[href*='/in/']",
            )

            if not links:
                return

            href = clean(links[0].get_attribute("href"))
            url = _url({"linkedin_url": href})

            if not url:
                return

            key = url.lower().split("?")[0].rstrip("/")

            if key in seen_urls:
                return

            seen_urls.add(key)
            found.append(card)

        except (
            WebDriverException,
            StaleElementReferenceException,
        ):
            pass

    # First: known LinkedIn result containers.
    for _ in range(5):
        for selector in selectors:
            try:
                cards = driver.find_elements(
                    By.CSS_SELECTOR,
                    selector,
                )
            except WebDriverException:
                continue

            for card in cards:
                add_card(card)

        try:
            driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);"
            )
        except WebDriverException:
            pass

        time.sleep(1.5)

    # Fallback: collect profile links directly.
    # This is deliberately broad because LinkedIn may change the
    # result-card wrapper while keeping /in/ profile URLs.
    try:
        profile_links = driver.find_elements(
            By.CSS_SELECTOR,
            "a[href*='/in/']",
        )

        for link in profile_links:
            try:
                href = clean(link.get_attribute("href"))
                url = _url({"linkedin_url": href})

                if not url:
                    continue

                key = url.lower().split("?")[0].rstrip("/")

                if key in seen_urls:
                    continue

                # Walk upward until we find a useful container.
                card = None
                current = link

                for _ in range(8):
                    try:
                        current = current.find_element(
                            By.XPATH,
                            "..",
                        )
                    except WebDriverException:
                        break

                    try:
                        text = clean(current.text)

                        if len(text) >= 20:
                            card = current
                            break
                    except WebDriverException:
                        continue

                if card is not None:
                    seen_urls.add(key)
                    found.append(card)

            except (
                WebDriverException,
                StaleElementReferenceException,
            ):
                continue

    except WebDriverException:
        pass

    print(
        f"People result containers/profile cards collected: {len(found)}"
    )

    return found

def _open_company_people_page(driver, company_url, company):
    """
    Open the company's LinkedIn page first, then move to its People section.

    We deliberately use the normal LinkedIn company UI rather than the
    global people search. If the People link is not exposed in the page,
    the standard /people/ company route is used as a fallback.
    """
    company_url = clean(company_url).split("?")[0].rstrip("/")

    if not company_url:
        return False

    try:
        print(f"Opening company page: {company_url}")
        driver.get(company_url)
        time.sleep(PEOPLE_WAIT_SECONDS)
    except WebDriverException:
        return False

    # First try the actual People tab/link shown on the company page.
    people_selectors = (
        "a[href*='/people/']",
        "a[href*='people']",
    )

    for selector in people_selectors:
        try:
            links = driver.find_elements(By.CSS_SELECTOR, selector)
        except WebDriverException:
            continue

        for link in links:
            try:
                href = clean(link.get_attribute("href"))
                text = clean(link.text).lower()

                if "/people/" not in href and "people" not in text:
                    continue

                if href:
                    print("Opening company People section:", href)
                    driver.get(href)
                    time.sleep(PEOPLE_WAIT_SECONDS)
                    return True
            except (WebDriverException, StaleElementReferenceException):
                continue

    # Fallback for layouts where LinkedIn does not expose the tab as a
    # normal anchor in the rendered DOM.
    people_url = f"{company_url}/people/"

    try:
        print("People tab link not exposed; opening company People route:", people_url)
        driver.get(people_url)
        time.sleep(PEOPLE_WAIT_SECONDS)
        return True
    except WebDriverException:
        return False


def search_people_for_job(
    driver,
    job,
    max_people=MAX_PEOPLE_PER_JOB,
):
    """
    Discover referral contacts from the COMPANY'S LINKEDIN PEOPLE SECTION.

    Flow:
        saved job -> Company URL -> LinkedIn company page -> People tab
        -> collect people cards -> filter relevant roles -> active people

    This intentionally does NOT use LinkedIn's global people search.
    """
    job_id = clean(job.get("job_id", job.get("Job ID", "")))
    company = clean(job.get("company", job.get("Company", "")))
    company_url = clean(job.get("company_url", job.get("Company URL", "")))

    if not job_id or not company:
        return []

    if not company_url:
        print(f"No Company URL stored for {company}; cannot open company People section.")
        return []

    print(f"\nSearching company People section for: {company}")

    if not _open_company_people_page(driver, company_url, company):
        print(f"Could not open People section for {company}.")
        return []

    # Make sure the page really looks like the company People page before
    # extracting cards. This also gives LinkedIn's SPA a little time to settle.
    try:
        current_url = clean(driver.current_url)
        print("People page URL:", current_url)
    except WebDriverException:
        pass

    cards = _collect_people_cards(driver)

    contacts = []
    seen = set()

    for card in cards:
        if len(contacts) >= max_people:
            break

        try:
            data = _extract_card(card)
            if not data:
                continue

            profile_url = data["linkedin_url"].lower()
            if profile_url in seen:
                continue
            seen.add(profile_url)

            # Company People pages can contain many employees. Keep the
            # configured referral roles plus common title variants that LinkedIn
            # may render differently (for example "Engineering Lead").
            role = clean(data.get("role", ""))
            role_low = role.casefold()

            role_match = any(
                role_name.casefold() in role_low
                for role_name in PEOPLE_ROLES
            ) or any(
                marker in role_low
                for marker in (
                    "recruiter",
                    "recruiting",
                    "talent acquisition",
                    "talent partner",
                    "hiring manager",
                    "engineering manager",
                    "engineering lead",
                    "technical lead",
                    "tech lead",
                    "principal engineer",
                    "staff engineer",
                    "lead engineer",
                    "senior software engineer",
                    "software engineer",
                    "software developer",
                    "developer",
                    "engineer",
                )
            )

            if not role or not role_match:
                continue

            # No visible recent activity = no outreach candidate.
            if not is_active_contact(data):
                continue

            data.update(
                {
                    "job_id": job_id,
                    "company": company,
                    "company_url": company_url,
                    "contact_status": "Discovered",
                    "request_status": "Not Requested",
                    "response_status": "No Response",
                    "referral_status": "Not Requested",
                }
            )

            data["contact_score"] = contact_score(data)
            data["referral_priority_score"] = referral_priority_score(data)
            data["high_value_role"] = (
                "Yes" if is_high_value_role(data["role"]) else "No"
            )

            # Save every discovered candidate through the existing storage
            # path so the job/person relationship stays job-specific.
            try:
                upsert_contact(data)
            except Exception:
                pass

            contacts.append(data)

            print(
                f"  FOUND: {data['person']} | {data['role']} | "
                f"connection={data.get('connection_degree', '') or 'unknown'} | "
                f"activity={data.get('activity', '') or 'none'} | "
                f"priority={data['referral_priority_score']}"
            )

        except (
            WebDriverException,
            StaleElementReferenceException,
        ):
            continue

    # Designation is dominant, then connection, then recent activity.
    contacts.sort(
        key=lambda x: (
            x["referral_priority_score"],
            x["contact_score"],
            -(
                activity_age_days(
                    x.get("activity", "")
                )
                if activity_age_days(x.get("activity", "")) is not None
                else 9999
            ),
        ),
        reverse=True,
    )

    print(f"People discovered from company People section: {len(contacts)}")
    return contacts[:max_people]


def _default_budget():
    return {
        "connection_requests": 0,
        "referral_messages": 0,
    }


def _can_send_connection(budget):
    return (
        budget["connection_requests"]
        < MAX_CONNECTION_REQUESTS_PER_RUN
    )


def _can_send_referral(budget):
    return (
        budget["referral_messages"]
        < MAX_REFERRAL_MESSAGES_PER_RUN
    )


def _get_job_link(job):
    """Return the saved job URL, or derive it from the LinkedIn Job ID."""
    job_url = clean(job.get("Job URL", job.get("job_url", "")))
    job_id = clean(job.get("Job ID", job.get("job_id", "")))

    if job_url:
        return job_url

    if job_id and job_id.isdigit():
        return f"https://www.linkedin.com/jobs/view/{job_id}/"

    return ""

def _referral_message_already_sent(job_id, person):
    """Return True only when this exact job/person referral message is recorded as Sent."""
    try:
        messages = _read_sheet("Messages")

        if messages is None or messages.empty:
            return False

        required = {
            "Job ID",
            "Person",
            "Message Type",
            "Status",
        }
        if not required.issubset(set(messages.columns)):
            return False

        rows = messages[
            (messages["Job ID"].astype(str).str.strip() == str(job_id).strip())
            & (
                messages["Person"].astype(str).str.strip().str.lower()
                == str(person).strip().lower()
            )
            & (
                messages["Message Type"].astype(str).str.strip().str.lower()
                == "referral request"
            )
            & (
                messages["Status"].astype(str).str.strip().str.lower()
                == "sent"
            )
        ]

        return not rows.empty

    except Exception as exc:
        # Do not block a referral merely because the spreadsheet could not be
        # read. The contact-state check below remains a second safeguard.
        print(f"Storage warning (check existing referral): {exc}")
        return False


def _send_referral_for_contact(driver, job, contact, budget):
    """Send the accepted-connection referral exactly once for Job ID + Person.

    Safety model:
      1. Check persistent local lock before opening the profile.
      2. Check Messages sheet/contact state.
      3. Open Message and wait for the real composer.
      4. Claim persistent lock BEFORE typing.
      5. Type + attach resume + click Send.
      6. Persist "sent" immediately after the Send click.
      7. Never remove the lock, even if later Excel updates fail.
    """
    if not _can_send_referral(budget):
        return False

    job_id = clean(job.get("Job ID", job.get("job_id", "")))
    company = clean(job.get("Company", job.get("company", "")))
    title = clean(job.get("Job Title", job.get("title", "")))
    person = _person(contact)
    profile_url = _url(contact)
    resume_path = clean(job.get("Resume Path", job.get("resume_path", "")))
    job_url = _get_job_link(job)

    if not person or not profile_url:
        return False

    referral_key = (str(job_id).strip(), str(person).strip().casefold())

    if referral_key in budget["_sent_referral_keys"]:
        print(
            f"Skipping duplicate referral in this run for {person} | job={job_id}."
        )
        return False

    # SAFETY CHECK 1: persistent local lock, before opening/clicking Message.
    lock_state = _get_referral_send_lock(job_id, person)
    if lock_state == "__SAFETY_STATE_UNREADABLE__":
        print(
            f"BLOCKED referral for {person}: persistent safety state could "
            "not be verified."
        )
        return False
    if lock_state:
        print(
            f"Skipping referral for {person}: persistent referral lock exists "
            f"for Job ID {job_id} (state={lock_state.get('state', 'unknown')})."
        )
        return False

    # SAFETY CHECK 2: exact Messages-sheet record.
    if _referral_message_already_sent(job_id, person):
        print(
            f"Skipping referral for {person}: referral message is already "
            f"recorded as Sent for Job ID {job_id}."
        )
        # Preserve that fact independently of Excel.
        if _claim_referral_send_lock(job_id, person):
            _mark_referral_send_lock(
                job_id, person, "sent", source="Messages sheet"
            )
        return False

    response_status = clean(
        contact.get("Response Status", contact.get("response_status", ""))
    ).lower()
    referral_status = clean(
        contact.get("Referral Status", contact.get("referral_status", ""))
    ).lower()

    if referral_status in {
        "requested", "sent", "offered", "received", "declined",
        "alternative process",
    }:
        print(
            f"Skipping referral for {person}: referral status is already "
            f"{referral_status}."
        )
        return False

    request_status_lower = clean(
        contact.get("Request Status", contact.get("request_status", ""))
    ).lower()

    if request_status_lower == "referral requested":
        print(
            f"Skipping referral for {person}: referral request is already "
            f"recorded for Job ID {job_id}."
        )
        return False

    if response_status not in {"", "no response"}:
        return False

    if not is_active_contact(contact):
        add_activity(
            job_id, company, person, "Referral Skipped - Inactive",
            f"Activity: {contact.get('Activity', contact.get('activity', ''))}",
        )
        return False

    if not resume_path:
        add_activity(
            job_id, company, person, "Referral Request Failed",
            "Selected Resume Path is empty; referral was not sent without the resume attachment.",
        )
        print(f"Referral blocked for {person}: Resume Path is empty.")
        return False

    if not os.path.isfile(resume_path):
        add_activity(
            job_id, company, person, "Referral Request Failed",
            f"Resume file was not found: {resume_path}",
        )
        print(f"Referral blocked for {person}: resume file not found: {resume_path}")
        return False

    if not job_url:
        add_activity(
            job_id, company, person, "Referral Request Failed",
            "No Job URL was available and a canonical URL could not be derived from Job ID.",
        )
        print(f"Referral blocked for {person}: job URL is missing.")
        return False

    message = referral_message(
        person,
        company,
        title,
        job_id,
        job_url,
        clean(job.get("Matched Skills", job.get("matched_skills", ""))).split(",")
        if job.get("Matched Skills", job.get("matched_skills", ""))
        else [],
    )

    if job_url not in message:
        message = f"{message.rstrip()}\n\nJob link: {job_url}"

    print(
        f"Preparing referral to {person} | job={job_id} | "
        f"resume={os.path.basename(resume_path)}"
    )

    try:
        # Open profile only after every pre-flight duplicate check.
        driver.get(profile_url)
        time.sleep(PEOPLE_WAIT_SECONDS)

        if not _click_profile_action(driver, "Message", timeout=12):
            print(
                f"Referral blocked for {person}: Message button was not "
                "exposed/clickable."
            )
            return False

        composer = None
        composer_deadline = time.time() + 20
        while time.time() < composer_deadline:
            composer = _find_message_composer(driver)
            if composer is not None:
                break
            time.sleep(0.35)

        if composer is None:
            print(f"Referral blocked for {person}: message composer was not exposed.")
            try:
                driver.save_screenshot("logs/message_composer_not_exposed.png")
            except WebDriverException:
                pass
            return False

        # CRITICAL: lock BEFORE typing. Ctrl+C during typing can therefore
        # never cause the next run to type the same message again.
        if not _claim_referral_send_lock(job_id, person):
            print(
                f"Referral blocked for {person}: could not claim the persistent "
                "pre-send safety lock."
            )
            return False

        print(
            f"Message composer exposed; referral send lock is active for "
            f"{person}. Entering referral message."
        )

        if not _clear_and_type_message(driver, composer, message):
            print(
                f"Referral text entry failed for {person}. Persistent lock "
                "remains active; this referral will NOT be auto-retried."
            )
            _mark_referral_send_lock(
                job_id, person, "started",
                note="Typing started but text entry did not complete.",
            )
            try:
                driver.save_screenshot("logs/message_text_entry_failed.png")
            except WebDriverException:
                pass
            return False

        time.sleep(0.8)

        # Resume attachment is mandatory.
        attached = _attach_resume_to_message(
            driver, resume_path, timeout=20
        )
        if not attached:
            print(
                f"Referral blocked for {person}: required resume attachment "
                "could not be verified. Persistent lock remains active; "
                "this referral will NOT be auto-retried."
            )
            _mark_referral_send_lock(
                job_id, person, "started",
                note="Message text entered but required attachment was not verified.",
            )
            try:
                driver.save_screenshot("logs/message_attachment_failed.png")
            except WebDriverException:
                pass
            return False

        # LinkedIn may render Send as text OR as an icon-only button.
        # The icon-only path is handled by _find_message_send_control().
        send_button = None
        send_deadline = time.time() + 15
        while time.time() < send_deadline:
            send_button = _find_message_send_control(driver)

            if send_button is not None:
                try:
                    if send_button.is_enabled():
                        break
                except (WebDriverException, StaleElementReferenceException):
                    pass

            send_button = None
            time.sleep(0.35)

        if send_button is None:
            print(
                f"Referral Send control was not exposed/enabled for {person}. "
                "Persistent lock remains active; this referral will NOT be "
                "auto-retried."
            )
            _mark_referral_send_lock(
                job_id, person, "started",
                note="Message typed and attachment verified, but Send was unavailable.",
            )
            try:
                driver.save_screenshot("logs/message_send_not_exposed.png")
            except WebDriverException:
                pass
            return False

        print("FOUND MESSAGE SEND CONTROL (text or icon)")
        print(
            f"CLICKING MESSAGE SEND: {person} | job={job_id}. "
            "This referral is one-time-only."
        )

        if not _click_visible_action(send_button, "Message Send"):
            print(
                f"Could not click Send for {person}. Persistent lock remains "
                "active; this referral will NOT be auto-retried."
            )
            _mark_referral_send_lock(
                job_id, person, "started",
                note="Send control was found but click did not complete.",
            )
            return False

        # IMPORTANT: the UI Send click is now considered irreversible.
        # Persist this immediately, before any other storage call.
        _mark_referral_send_lock(
            job_id, person, "sent",
            sent_at=time.strftime("%Y-%m-%d %H:%M:%S"),
            job_url=job_url,
            resume_name=os.path.basename(resume_path),
        )

        time.sleep(2)

        print(
            "Message send click completed with resume attachment. "
            f"Referral permanently locked for {person} | job={job_id}."
        )

        # Excel is secondary persistence now; failure here cannot cause retry.
        _safe_storage_call(
            "add referral message record",
            add_message,
            job_id=job_id,
            company=company,
            person=person,
            message_type="Referral Request",
            message=message,
            resume_type=job.get("Resume Type", ""),
            resume_name=job.get("Resume Name", ""),
            resume_path=resume_path,
            job_url=job_url,
            attachment_required="Yes",
            status="Sent",
        )

        budget["referral_messages"] += 1
        budget["_sent_referral_keys"].add(referral_key)

        _safe_storage_call(
            "referral message sent",
            mark_message_sent,
            job_id, person, "Referral Request",
        )
        _safe_storage_call(
            "referral request recorded",
            mark_referral_requested,
            job_id, person,
        )

        add_activity(
            job_id,
            company,
            person,
            "Referral Request Sent",
            "Sent after connection acceptance with job link and resume attachment: "
            f"{os.path.basename(resume_path)}",
        )
        return True

    except (WebDriverException, StaleElementReferenceException) as exc:
        print(
            f"LinkedIn referral UI failed for {person}: {exc}. "
            "Persistent lock remains if claimed; no automatic retry."
        )
        lock = _get_referral_send_lock(job_id, person)
        if lock not in {None, "__SAFETY_STATE_UNREADABLE__"}:
            _mark_referral_send_lock(
                job_id, person, "started",
                note=f"Browser exception after referral lock was claimed: {exc}",
            )
        return False

    except Exception as exc:
        print(
            f"LinkedIn referral failed for {person}: {exc}. "
            "Persistent lock remains if claimed; no automatic retry."
        )
        lock = _get_referral_send_lock(job_id, person)
        if lock not in {None, "__SAFETY_STATE_UNREADABLE__"}:
            _mark_referral_send_lock(
                job_id, person, "started",
                note=f"Exception after referral lock was claimed: {exc}",
            )
        return False

def _send_connection_for_discovered_contact(driver, job, contact, budget):
    """Send a connection request for a newly discovered saved-job contact only.

    IMPORTANT: this function NEVER sends the referral request while sending the
    connection invitation. The referral message is a separate action that can
    happen only after LinkedIn later shows the connection as accepted.
    """
    job_id = clean(job.get("Job ID", job.get("job_id", "")))
    company = clean(job.get("Company", job.get("company", "")))
    title = clean(job.get("Job Title", job.get("title", "")))
    person = _person(contact)
    profile_url = _url(contact)

    if not job_id or not person or not profile_url:
        return False

    if not _can_send_connection(budget):
        print(f"Connection request deferred for {person}: per-run connection limit reached.")
        return False

    # Never send another connection request to a person already contacted for
    # this exact job.
    existing = next(
        (
            item
            for item in get_contacts_for_job(job_id)
            if clean(item.get("Person", "")).casefold() == person.casefold()
        ),
        {},
    )
    request_status = clean(existing.get("Request Status", "")).lower()
    if request_status in {"connection requested", "requested", "accepted"}:
        return False

    connection_note = _connection_note_for_contact(
        person, company, title, contact
    )

    # This is deliberately a CONNECTION-ONLY operation. Do not pass referral
    # text into send_connection_request.
    sent = send_connection_request(
        driver,
        profile_url,
        note=connection_note or None,
        message_text=None,
    )

    if not sent:
        return False

    budget["connection_requests"] += 1

    merged = dict(contact)
    merged.update(existing)
    merged["Job ID"] = job_id
    merged["Person"] = person
    merged["Request Status"] = "Connection Requested"
    merged["Contact Status"] = "Contacted"
    merged["Response Status"] = merged.get("Response Status") or "No Response"
    merged["Referral Status"] = merged.get("Referral Status") or "Not Requested"

    try:
        upsert_contact(merged)
    except Exception as exc:
        print(f"Storage warning (connection contact {person}): {exc}")

    _safe_storage_call(
        "connection request message record",
        add_message,
        job_id=job_id,
        company=company,
        person=person,
        message_type="Connection Request",
        message=(
            "Sent with a personalized connection note; referral message will be "
            "sent only after connection acceptance."
            if connection_note
            else "Sent without a note; referral message will be sent only after connection acceptance."
        ),
        resume_type=job.get("Resume Type", ""),
        resume_name=job.get("Resume Name", ""),
        resume_path=job.get("Resume Path", ""),
        job_url=_get_job_link(job),
        attachment_required="No",
        status="Sent",
    )
    _safe_storage_call(
        "connection request status",
        mark_message_sent,
        job_id,
        person,
        "Connection Request",
    )
    add_activity(
        job_id,
        company,
        person,
        "Connection Request Sent",
        "Saved-job top-up contact. Referral request will be sent only after acceptance.",
    )

    print(
        f"Connection top-up sent to {person} | job={job_id} | "
        f"company={company}"
    )
    return True


def _top_up_saved_job_contacts(driver, job, budget):
    """Bring an older saved job up to the configured contact target.

    This is an ACTION stage, not a discovery-only stage.

    For every newly discovered active candidate:
      * connected -> immediately send referral message + resume
      * connect_available -> send connection request only
      * pending -> do nothing and keep looking for another candidate

    Duplicate protection is Job ID + Person. One person's referral does not
    block another person's referral for the same job.
    """
    job_id = clean(job.get("Job ID", job.get("job_id", "")))
    if not job_id:
        return 0

    existing_contacts = get_contacts_for_job(job_id)
    contacted_people = set()

    for item in existing_contacts:
        person = clean(item.get("Person", "")).casefold()
        if not person:
            continue

        request_status = clean(item.get("Request Status", "")).lower()
        contact_status = clean(item.get("Contact Status", "")).lower()
        referral_status = clean(item.get("Referral Status", "")).lower()

        if (
            request_status in {"connection requested", "requested", "accepted", "referral requested"}
            or contact_status in {"contacted", "connected"}
            or referral_status in {
                "requested", "sent", "offered", "received", "declined",
                "alternative process",
            }
        ):
            contacted_people.add(person)

    current_count = len(contacted_people)
    target = MAX_INITIAL_REFERRAL_CONTACTS

    if current_count >= target:
        return 0

    remaining = target - current_count
    print(
        f"Saved job {job_id} has {current_count}/{target} contacted people. "
        f"Searching for up to {remaining} additional candidates."
    )

    # Search beyond the exact shortfall because candidates may be pending,
    # already contacted, inactive, or otherwise unavailable.
    discovery_limit = max(MAX_PEOPLE_PER_JOB, target + 5)
    contacts = search_people_for_job(
        driver,
        job,
        max_people=discovery_limit,
    )

    added = 0

    for contact in contacts:
        if len(contacted_people) >= target:
            break

        person = _person(contact)
        profile_url = _url(contact)
        key = person.casefold()

        if not person or not profile_url or key in contacted_people:
            continue

        # Exact Job ID + Person duplicate protection.
        lock_state = _get_referral_send_lock(job_id, person)
        if lock_state == "__SAFETY_STATE_UNREADABLE__":
            print(
                f"Skipping {person}: persistent referral safety state could not be verified."
            )
            continue
        if lock_state:
            print(
                f"Skipping {person}: persistent referral lock exists "
                f"(state={lock_state.get('state', 'unknown')})."
            )
            continue
        if _referral_message_already_sent(job_id, person):
            print(
                f"Skipping {person}: referral message already recorded as Sent "
                f"for Job ID {job_id}."
            )
            continue

        existing = next(
            (
                item
                for item in get_contacts_for_job(job_id)
                if clean(item.get("Person", "")).casefold() == key
            ),
            {},
        )

        request_status = clean(existing.get("Request Status", "")).lower()
        response_status = clean(existing.get("Response Status", "")).lower()
        referral_status = clean(existing.get("Referral Status", "")).lower()

        if response_status not in {"", "no response"}:
            continue
        if referral_status in {
            "requested", "sent", "offered", "received", "declined",
            "alternative process",
        }:
            contacted_people.add(key)
            continue
        if request_status in {"connection requested", "requested", "accepted", "referral requested"}:
            contacted_people.add(key)
            continue

        # We have NOT contacted this person for this job yet. Inspect their
        # current LinkedIn UI and immediately perform the appropriate action.
        state = get_profile_connection_state(driver, profile_url)
        print(
            f"Top-up candidate: {person} | live connection state={state}"
        )

        if state == "pending":
            # This invitation was not sent by this workflow, so do not pretend
            # it was our connection request. Leave it untouched and look for a
            # different candidate.
            add_activity(
                job_id,
                job.get("Company", ""),
                person,
                "Top-Up Candidate Already Pending",
                "LinkedIn shows a pending invitation; no action taken.",
            )
            continue

        if state == "connected":
            # IMPORTANT FIX: a newly discovered connected person must receive
            # the referral immediately. Previously this branch only saved
            # Accepted/1st and then stopped, which caused the exact
            # "found people but sent nothing" behavior.
            merged = dict(contact)
            merged.update(existing)
            merged["Job ID"] = job_id
            merged["Person"] = person
            merged["Connection Degree"] = "1st"
            merged["Request Status"] = "Accepted"
            merged["Contact Status"] = "Connected"
            merged["Response Status"] = merged.get("Response Status") or "No Response"
            merged["Referral Status"] = merged.get("Referral Status") or "Not Requested"

            _safe_storage_call(
                "save newly discovered connected contact",
                upsert_contact,
                merged,
            )
            _safe_storage_call(
                "mark newly discovered connection accepted",
                mark_connection_accepted,
                job_id,
                person,
            )

            if _send_referral_for_contact(driver, job, merged, budget):
                contacted_people.add(key)
                added += 1
                print(
                    f"Top-up referral SENT to {person} | job={job_id}"
                )
            else:
                print(
                    f"Top-up referral was not sent to {person}; continuing to next candidate."
                )
            continue

        if state != "connect_available":
            add_activity(
                job_id,
                job.get("Company", ""),
                person,
                "Top-Up Candidate Skipped",
                f"LinkedIn profile state: {state}",
            )
            continue

        # Only connection requests consume the connection-request budget.
        if not _can_send_connection(budget):
            print(
                f"Cannot send connection to {person}: per-run connection limit reached."
            )
            continue

        if _send_connection_for_discovered_contact(driver, job, contact, budget):
            contacted_people.add(key)
            added += 1

    # Re-read the tracker for the authoritative saved count. Also report the
    # number of successful UI actions separately so a storage warning cannot
    # silently turn into a false 3/3 claim.
    final_saved_people = _contacted_people_for_job(job_id)
    final_saved_count = len(final_saved_people)
    print(
        f"Top-up complete for {job_id}: saved contacted people "
        f"{final_saved_count}/{target}; successful new actions this top-up={added}."
    )
    if final_saved_count < target and added:
        print(
            f"Top-up note for {job_id}: {added} successful action(s) were "
            f"performed, but only {final_saved_count}/{target} are currently "
            "persisted as contacted in the tracker."
        )
    return added


def process_pending_connections_for_job(
    driver,
    job,
    budget=None,
):
    """Continue saved workflows AND top up older saved jobs to the contact target.

    Rules:
      - Existing accepted connections may receive exactly one referral request.
      - Pending invitations receive no referral request.
      - Older jobs with only 1-2 contacted people are searched again and topped
        up to MAX_INITIAL_REFERRAL_CONTACTS, subject to the global run budget.
      - Sending a connection request NEVER sends the referral message.
      - Referral duplicate protection remains Job ID + Person.
    """
    budget = budget or _default_budget()

    if "_sent_referral_keys" not in budget:
        budget["_sent_referral_keys"] = set()

    job_id = clean(job.get("Job ID", ""))
    if not job_id:
        return 0

    actions = 0
    contacts = get_contacts_for_job(job_id)

    # First process people already saved in the Contacts sheet.
    for contact in contacts:
        person = _person(contact)
        profile_url = _url(contact)

        if not person or not profile_url:
            continue

        lock_state = _get_referral_send_lock(job_id, person)
        if lock_state == "__SAFETY_STATE_UNREADABLE__":
            print(
                f"Skipping referral for {person}: persistent safety state "
                "could not be verified."
            )
            continue
        if lock_state:
            print(
                f"Skipping referral for {person}: persistent referral lock "
                f"exists (state={lock_state.get('state', 'unknown')})."
            )
            continue
        if _referral_message_already_sent(job_id, person):
            print(
                f"Skipping referral for {person}: referral message is already "
                f"recorded as Sent for Job ID {job_id}."
            )
            if _claim_referral_send_lock(job_id, person):
                _mark_referral_send_lock(
                    job_id, person, "sent", source="Messages sheet"
                )
            continue

        request_status = clean(contact.get("Request Status", "")).lower()
        referral_status = clean(contact.get("Referral Status", "")).lower()
        response_status = clean(contact.get("Response Status", "")).lower()

        if response_status not in {"", "no response"}:
            continue
        if referral_status in {
            "requested", "sent", "offered", "received", "declined",
            "alternative process",
        }:
            continue

        known_workflow = (
            request_status in {"connection requested", "requested", "accepted"}
            or "1st" in clean(contact.get("Connection Degree", "")).lower()
        )
        if not known_workflow or not is_active_contact(contact):
            continue

        # Always inspect live LinkedIn state; spreadsheet state can be stale.
        state = get_profile_connection_state(driver, profile_url)

        if state == "pending":
            add_activity(
                job_id, job.get("Company", ""), person,
                "Connection Still Pending",
                "LinkedIn currently shows a pending invitation; no referral sent.",
            )
            print(f"Skipping referral for {person}: connection is still pending.")
            continue

        if state == "connected":
            mark_connection_accepted(job_id, person)
            merged = dict(contact)
            merged["Connection Degree"] = "1st"
            merged["Request Status"] = "Accepted"
            merged["Contact Status"] = "Connected"

            add_activity(
                job_id, job.get("Company", ""), person,
                "Connection Accepted",
                "Detected from current LinkedIn profile UI.",
            )

            if _send_referral_for_contact(driver, job, merged, budget):
                actions += 1
            continue

        add_activity(
            job_id, job.get("Company", ""), person,
            "Connection State Needs Review",
            f"LinkedIn profile state: {state}; no referral sent.",
        )

    # NEW: after continuing saved contacts, top up old saved jobs that have
    # fewer than the configured 3-4 contact target. This is why a job with only
    # 1-2 previous connection requests was previously stuck at 1-2 forever.
    try:
        actions += _top_up_saved_job_contacts(driver, job, budget)
    except Exception as exc:
        print(
            f"Saved-job contact top-up failed for Job ID {job_id}: {exc}"
        )

    return actions

def _connection_note_for_contact(person, company, job_title, contact):
    """Return a short note only for genuinely high-value contacts."""
    role = clean(contact.get("role", contact.get("Role", ""))).lower()
    high_value = (
        "ceo" in role
        or "chief executive" in role
        or "founder" in role
        or "co-founder" in role
        or "vp " in role
        or role.startswith("vp")
        or "vice president" in role
        or "director" in role
        or "head of" in role
        or is_high_value_role(role)
    )

    if not high_value:
        return ""

    return (
        f"Hi {person},\n\n"
        f"I came across your profile while exploring opportunities at {company}. "
        f"I'm currently interested in the {job_title} opportunity and would be glad to connect.\n\n"
        "Thanks!"
    )


def _contacted_people_for_job(job_id):
    """Return distinct people who have a real outreach workflow for this Job ID.

    A discovered/saved contact with ``Not Requested`` is NOT counted.
    A person counts only after a connection/referral workflow has actually been
    recorded in the Contacts sheet.
    """
    contacted = set()
    for item in get_contacts_for_job(job_id):
        person = _person(item)
        if not person:
            continue

        request_status = clean(item.get("Request Status", "")).casefold()
        contact_status = clean(item.get("Contact Status", "")).casefold()
        referral_status = clean(item.get("Referral Status", "")).casefold()

        if (
            request_status in {
                "connection requested",
                "requested",
                "accepted",
                "referral requested",
            }
            or contact_status in {"contacted", "connected"}
            or referral_status in {
                "requested",
                "sent",
                "offered",
                "received",
                "declined",
                "alternative process",
            }
        ):
            contacted.add(person.casefold())

    return contacted


def _log_job_contact_progress(job_id, target, prefix="Job contact progress"):
    """Log the saved, distinct contacted-person count for this exact Job ID."""
    contacted = _contacted_people_for_job(job_id)
    print(
        f"{prefix} for {job_id}: {len(contacted)}/{target} contacted people. "
        f"New connection/referral actions are counted separately."
    )
    return contacted


def process_people_for_job(
    driver,
    job,
    budget=None,
):
    """Process people for one job without falsely satisfying the contact target.

    Important counting rule:
      * ``MAX_INITIAL_REFERRAL_CONTACTS`` is a per-Job-ID target.
      * A person counts toward that target only after a real connection request
        or referral workflow has been recorded for that exact Job ID + Person.
      * Merely discovering a person, seeing a pending invitation, or opening a
        profile does NOT count.
      * The loop scans ALL discovered candidates (up to MAX_PEOPLE_PER_JOB) and
        does not stop after the first three unusable candidates.
      * Connection requests remain globally capped by
        MAX_CONNECTION_REQUESTS_PER_RUN.
    """
    budget = budget or _default_budget()

    if "_sent_referral_keys" not in budget:
        budget["_sent_referral_keys"] = set()

    job_id = clean(job.get("job_id", job.get("Job ID", "")))
    if not job_id:
        return []

    try:
        score = float(
            job.get(
                "Relevance Score",
                job.get("relevance_score", 0),
            )
            or 0
        )
    except (TypeError, ValueError):
        score = 0

    if score < MIN_RELEVANCE_SCORE:
        _safe_storage_call(
            "relevance-gate activity",
            add_activity,
            job_id,
            job.get("Company", ""),
            "",
            "Outreach Skipped - Relevance Below Threshold",
            f"Relevance Score={score}; threshold={MIN_RELEVANCE_SCORE}",
        )
        return []

    target = MAX_INITIAL_REFERRAL_CONTACTS
    before = _contacted_people_for_job(job_id)
    print(
        f"Job {job_id} contact target before processing: "
        f"{len(before)}/{target}."
    )

    # Continue saved pending/accepted workflows first. This can increase the
    # contacted count by sending referrals to accepted connections, but it does
    # not manufacture contact counts for skipped candidates.
    try:
        saved_actions = process_pending_connections_for_job(
            driver,
            job,
            budget,
        )
    except Exception as exc:
        saved_actions = 0
        print(
            f"Saved workflow processing failed for {job_id}: {exc}"
        )

    actions = []
    if saved_actions:
        actions.append({
            "person": "",
            "action": "Saved Workflow",
            "status": f"{saved_actions} action(s)",
        })

    current = _contacted_people_for_job(job_id)
    if len(current) >= target:
        print(
            f"Job {job_id} already has {len(current)}/{target} contacted people "
            "after saved-workflow processing; no new contacts needed."
        )
        return actions

    contacts = search_people_for_job(
        driver,
        job,
        max_people=MAX_PEOPLE_PER_JOB,
    )

    if not contacts:
        print(
            f"No eligible people discovered for Job ID {job_id}. "
            f"Contacted remains {len(current)}/{target}."
        )
        return actions

    # IMPORTANT: do not slice to the target here. The first few candidates may
    # be pending, locked, already contacted, or otherwise unusable. We must
    # continue through the complete discovered candidate pool until the per-job
    # target is genuinely reached or all candidates are exhausted.
    eligible = [
        contact
        for contact in contacts
        if is_active_contact(contact)
    ]

    new_actions = 0

    for contact in eligible:
        current = _contacted_people_for_job(job_id)
        if len(current) >= target:
            break

        person = _person(contact)
        profile_url = _url(contact)
        if not person or not profile_url:
            continue

        person_key = person.casefold()

        # If this person has already been genuinely contacted for this Job ID,
        # skip them but continue scanning later candidates.
        if person_key in current:
            print(
                f"Skipping {person}: already contacted for Job ID {job_id}; "
                "continuing to next candidate."
            )
            continue

        lock_state = _get_referral_send_lock(job_id, person)
        if lock_state == "__SAFETY_STATE_UNREADABLE__":
            print(
                f"Skipping {person}: persistent referral safety state "
                "could not be verified."
            )
            continue
        if lock_state:
            print(
                f"Skipping {person}: persistent referral lock exists "
                f"(state={lock_state.get('state', 'unknown')})."
            )
            continue

        if _referral_message_already_sent(job_id, person):
            print(
                f"Skipping {person}: referral message is already recorded "
                f"as Sent for Job ID {job_id}."
            )
            if _claim_referral_send_lock(job_id, person):
                _mark_referral_send_lock(
                    job_id,
                    person,
                    "sent",
                    source="Messages sheet",
                )
            continue

        # Refresh the exact Job ID + Person record because search_people_for_job
        # may have just discovered/upserted this person.
        existing = next(
            (
                item
                for item in get_contacts_for_job(job_id)
                if _person(item).casefold() == person_key
            ),
            {},
        )

        request_status = clean(
            existing.get(
                "Request Status",
                contact.get("request_status", ""),
            )
        ).casefold()
        response_status = clean(
            existing.get(
                "Response Status",
                contact.get("response_status", ""),
            )
        ).casefold()
        referral_status = clean(
            existing.get(
                "Referral Status",
                contact.get("referral_status", ""),
            )
        ).casefold()

        if response_status not in {"", "no response"}:
            print(
                f"Skipping {person}: they have a recorded response "
                f"({response_status})."
            )
            continue

        if referral_status in {
            "requested",
            "sent",
            "offered",
            "received",
            "declined",
            "alternative process",
        }:
            print(
                f"Skipping {person}: referral status is already "
                f"{referral_status}."
            )
            continue

        # Existing accepted connection: referral workflow only.
        if request_status == "accepted" or "1st" in clean(
            existing.get(
                "Connection Degree",
                contact.get("connection_degree", ""),
            )
        ).casefold():
            merged = dict(contact)
            merged.update(existing)
            merged["Job ID"] = job_id
            merged["Person"] = person
            merged["Connection Degree"] = "1st"
            merged["Request Status"] = "Accepted"
            merged["Contact Status"] = "Connected"
            merged["Response Status"] = (
                merged.get("Response Status") or "No Response"
            )
            merged["Referral Status"] = (
                merged.get("Referral Status") or "Not Requested"
            )

            if _send_referral_for_contact(
                driver,
                job,
                merged,
                budget,
            ):
                actions.append({
                    "person": person,
                    "action": "Referral Request",
                    "status": "Sent",
                })
                new_actions += 1
            continue

        # Existing pending/requested workflow is not a reason to stop looking
        # for other people. It also does not count as a NEW action in this loop.
        if request_status in {"connection requested", "requested"}:
            print(
                f"Skipping {person}: connection request already pending for "
                f"Job ID {job_id}; continuing to next candidate."
            )
            continue

        # Live LinkedIn state is authoritative over stale Excel state.
        state = get_profile_connection_state(driver, profile_url)
        print(
            f"Candidate {person} | live connection state={state} | "
            f"job={job_id}"
        )

        if state == "connected":
            mark_connection_accepted(job_id, person)

            merged = dict(contact)
            merged.update(existing)
            merged["Job ID"] = job_id
            merged["Person"] = person
            merged["Connection Degree"] = "1st"
            merged["Request Status"] = "Accepted"
            merged["Contact Status"] = "Connected"
            merged["Response Status"] = (
                merged.get("Response Status") or "No Response"
            )
            merged["Referral Status"] = (
                merged.get("Referral Status") or "Not Requested"
            )

            if _send_referral_for_contact(
                driver,
                job,
                merged,
                budget,
            ):
                actions.append({
                    "person": person,
                    "action": "Referral Request",
                    "status": "Sent",
                })
                new_actions += 1
            continue

        if state == "pending":
            add_activity(
                job_id,
                job.get("Company", ""),
                person,
                "Existing Connection Request Found",
                "LinkedIn shows a pending invitation; no duplicate invitation sent. Continuing to another candidate.",
            )
            print(
                f"Skipping {person}: LinkedIn invitation is pending; "
                "not counted as a new contact."
            )
            continue

        if state != "connect_available":
            add_activity(
                job_id,
                job.get("Company", ""),
                person,
                "Connection Skipped",
                f"LinkedIn profile state: {state}",
            )
            print(
                f"Skipping {person}: live LinkedIn state={state}; "
                "continuing to next candidate."
            )
            continue

        if not _can_send_connection(budget):
            add_activity(
                job_id,
                job.get("Company", ""),
                person,
                "Connection Request Deferred",
                "Global per-run connection-request limit reached.",
            )
            print(
                f"Stopping new connection outreach for {job_id}: global "
                f"limit {MAX_CONNECTION_REQUESTS_PER_RUN} reached."
            )
            break

        company = clean(job.get("Company", ""))
        title = clean(job.get("Job Title", ""))
        connection_note = _connection_note_for_contact(
            person,
            company,
            title,
            contact,
        )

        # INITIAL OUTREACH IS CONNECTION-ONLY.
        sent = send_connection_request(
            driver,
            profile_url,
            note=connection_note or None,
            message_text=None,
        )

        _safe_storage_call(
            "connection request message record",
            add_message,
            job_id=job_id,
            company=company,
            person=person,
            message_type="Connection Request",
            message=(
                "Sent with a personalized connection note; referral message will be sent after connection acceptance."
                if connection_note
                else "Sent without a note; referral message will be sent after connection acceptance."
            ),
            resume_type=job.get("Resume Type", ""),
            resume_name=job.get("Resume Name", ""),
            resume_path=job.get("Resume Path", ""),
            job_url=_get_job_link(job),
            attachment_required="No",
            status="Sent" if sent else "Failed",
        )

        if not sent:
            add_activity(
                job_id,
                company,
                person,
                "Connection Request Failed",
                "LinkedIn UI did not complete the send action; person does not count toward the per-job contact target.",
            )
            actions.append({
                "person": person,
                "action": "Connection Request",
                "status": "Failed",
            })
            time.sleep(2)
            continue

        # Increment only after send_connection_request reports success.
        budget["connection_requests"] += 1

        _safe_storage_call(
            "connection request sent marker",
            mark_message_sent,
            job_id,
            person,
            "Connection Request",
        )

        # Persist the person as genuinely contacted for this Job ID. This is
        # what makes the per-job 3/3 counter truthful on the next calculation.
        merged = dict(contact)
        merged.update(existing)
        merged["Job ID"] = job_id
        merged["Person"] = person
        merged["Request Status"] = "Connection Requested"
        merged["Contact Status"] = "Contacted"
        merged["Response Status"] = (
            merged.get("Response Status") or "No Response"
        )
        merged["Referral Status"] = (
            merged.get("Referral Status") or "Not Requested"
        )
        _safe_storage_call(
            "save contacted person after successful connection request",
            upsert_contact,
            merged,
        )

        add_activity(
            job_id,
            company,
            person,
            "Connection Request Sent",
            "Active contact selected. Connection request was successfully reported by the LinkedIn UI; referral remains blocked until acceptance.",
        )

        actions.append({
            "person": person,
            "action": "Connection Request",
            "status": "Sent",
        })
        new_actions += 1

        # Re-read saved state after every successful action. Do not infer 3/3
        # solely from an in-memory counter.
        saved_now = _contacted_people_for_job(job_id)
        print(
            f"Job {job_id} progress after contacting {person}: "
            f"{len(saved_now)}/{target} contacted people; "
            f"run connection requests={budget['connection_requests']}/{MAX_CONNECTION_REQUESTS_PER_RUN}."
        )

        time.sleep(2)

    final_contacted = _contacted_people_for_job(job_id)
    print(
        f"Final contact count for {job_id}: "
        f"{len(final_contacted)}/{target} contacted people. "
        f"New successful outreach actions this job={new_actions}. "
        f"Global connection requests this run={budget['connection_requests']}/{MAX_CONNECTION_REQUESTS_PER_RUN}."
    )

    return actions

def prepare_updates_after_referral(job, referrer_person):
    """
    Prepare updates only for contacts who explicitly offered to refer.
    """
    job_id = clean(
        job.get("job_id", job.get("Job ID", ""))
    )

    company = clean(
        job.get("company", job.get("Company", ""))
    )

    title = clean(
        job.get("title", job.get("Job Title", ""))
    )

    if not referral_already_received(job_id):
        return []

    prepared = []

    # Import lazily so this module remains compatible with the existing
    # storage layer.
    from storage.excel import get_willing_contacts_except

    for contact in get_willing_contacts_except(
        job_id,
        excluded_person=referrer_person,
    ):
        person = clean(
            contact.get("Person", "")
        )

        message = prepare_referral_already_received_message(
            person,
            company,
            title,
            job_id,
        )

        add_message(
            job_id=job_id,
            company=company,
            person=person,
            message_type="Referral Already Received Update",
            message=message,
            job_url=job.get("Job URL", ""),
            attachment_required="No",
            status="Prepared",
        )

        prepared.append(
            {
                "person": person,
                "message": message,
            }
        )

    return prepared


def send_updates_after_referral(
    driver,
    job,
    referrer_person,
):
    results = []

    for item in prepare_updates_after_referral(
        job,
        referrer_person,
    ):
        contact = next(
            (
                c
                for c in get_contacts_for_job(
                    job.get("job_id", job.get("Job ID", ""))
                )
                if clean(c.get("Person", "")).lower()
                == item["person"].lower()
            ),
            {},
        )

        sent = send_linkedin_message(
            driver,
            contact.get("LinkedIn URL", ""),
            item["message"],
        )

        results.append(
            {
                "person": item["person"],
                "status": "Sent" if sent else "Failed",
            }
        )

        if sent:
            mark_message_sent(
                job.get("job_id", job.get("Job ID", "")),
                item["person"],
                "Referral Already Received Update",
            )

    return results
def get_profile_connection_state(driver, profile_url):
    """Determine connection state from the normal rendered LinkedIn profile UI.

    Priority is intentional:
        pending/invitation sent -> pending
        Connect (even if Message is also visible) -> connect_available
        Message with no Connect/Pending -> connected
        otherwise -> unavailable

    For this persisted connection-request workflow, a pending invitation must
    never trigger a referral, even if LinkedIn also exposes Message.
    """
    if not profile_url:
        return "unavailable"

    try:
        driver.get(profile_url)
        time.sleep(PEOPLE_WAIT_SECONDS)

        elements = driver.find_elements(
            By.CSS_SELECTOR,
            "button, a, [role='button']",
        )

        visible_controls = []
        has_message = False
        has_connect = False
        has_pending = False

        for element in elements:
            try:
                if not element.is_displayed() or not element.is_enabled():
                    continue

                text = clean(element.text)
                aria = clean(element.get_attribute("aria-label"))
                title = clean(element.get_attribute("title"))
                combined = f"{text} {aria} {title}".strip().casefold()
                low_text = text.casefold()
                low_aria = aria.casefold()
                low_title = title.casefold()

                if combined:
                    visible_controls.append(combined)

                if (
                    low_text == "message"
                    or low_aria == "message"
                    or low_title == "message"
                ):
                    has_message = True

                if (
                    low_text == "connect"
                    or low_aria == "connect"
                    or low_title == "connect"
                    or ("invite" in low_aria and "connect" in low_aria)
                ):
                    has_connect = True

                if any(
                    marker in combined
                    for marker in (
                        "invitation sent",
                        "request sent",
                        "invitation pending",
                        "pending",
                    )
                ):
                    has_pending = True

            except (WebDriverException, StaleElementReferenceException):
                continue

        # HARD UI RULE:
        # 1) Pending always wins.
        # 2) Connect wins over Message when both are visible and Pending is absent.
        #    LinkedIn can expose Message + Connect together for profiles that are
        #    NOT connected. Message is therefore NOT proof of an accepted connection.
        # 3) Message with no Connect/Pending means connected.
        if has_pending:
            print(
                f"  DEBUG connection controls | Message={'YES' if has_message else 'NO'} "
                f"| Connect={'YES' if has_connect else 'NO'} | Pending=YES"
            )
            return "pending"

        if has_connect:
            print(
                f"  DEBUG connection controls | Message={'YES' if has_message else 'NO'} "
                "| Connect=YES | Pending=NO -> CONNECT_AVAILABLE"
            )
            return "connect_available"

        if has_message:
            print(
                "  DEBUG connection controls | Message=YES | Connect=NO | Pending=NO "
                "-> CONNECTED"
            )
            return "connected"

        # Body-level fallback is only used for strong pending markers.  Do not
        # use a generic 'message' substring because job/page text can contain it.
        page_text = ""
        try:
            page_text = clean(
                driver.find_element(By.TAG_NAME, "body").text
            ).casefold()
        except WebDriverException:
            pass

        if any(
            marker in page_text
            for marker in (
                "invitation sent",
                "request sent",
                "invitation pending",
            )
        ):
            return "pending"

        return "unavailable"

    except (WebDriverException, StaleElementReferenceException):
        return "unavailable"
    except Exception:
        return "unavailable"

def _find_message_attachment_control(driver):
    """Find a visible attachment control in the open LinkedIn message composer."""
    # Search shadow DOM first; LinkedIn's composer controls can live under
    # #interop-outlet and are invisible to normal Selenium CSS queries.
    shadow = _find_message_action_in_shadow(driver, "Attach")
    if shadow is not None:
        return shadow
    for candidate in ("Attach media", "Add a file", "Attach file", "Attachment"):
        shadow = _find_message_action_in_shadow(driver, candidate)
        if shadow is not None:
            return shadow
    selectors = "button, [role='button'], a"
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selectors)
    except WebDriverException:
        return None

    markers = (
        "attach",
        "attachment",
        "add a file",
        "add attachment",
        "add an attachment",
        "attach file",
        "add media",
    )

    for element in elements:
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue
            text = clean(element.text).casefold()
            aria = clean(element.get_attribute("aria-label")).casefold()
            title = clean(element.get_attribute("title")).casefold()
            combined = f"{text} {aria} {title}"
            if any(marker in combined for marker in markers):
                return element
        except (WebDriverException, StaleElementReferenceException):
            continue

    return None


def _find_message_file_input(driver):
    """Find the file input for the open LinkedIn message composer.

    Prefer the real HTML input so send_keys() can select the resume directly
    without opening the Linux/GTK file picker.
    """
    candidates = []

    try:
        candidates.extend(
            _find_shadow_dom_elements(
                driver,
                "input[type='file']",
                visible_only=False,
            ) or []
        )
    except WebDriverException:
        pass

    try:
        candidates.extend(
            driver.find_elements(
                By.CSS_SELECTOR,
                "input[type='file']",
            )
        )
    except WebDriverException:
        pass

    ranked = []
    seen = set()

    for element in candidates:
        try:
            key = element.id
        except Exception:
            key = str(id(element))

        if key in seen:
            continue

        seen.add(key)

        try:
            accept = clean(
                element.get_attribute("accept")
            ).casefold()
            name = clean(
                element.get_attribute("name")
            ).casefold()
            aria = clean(
                element.get_attribute("aria-label")
            ).casefold()
            cls = clean(
                element.get_attribute("class")
            ).casefold()

            combined = f"{accept} {name} {aria} {cls}"

            score = 0

            if "pdf" in accept or "application/pdf" in accept:
                score += 100

            if any(
                marker in combined
                for marker in (
                    "attachment",
                    "attach",
                    "document",
                    "file",
                    "message",
                    "messaging",
                )
            ):
                score += 50

            ranked.append((score, element))

        except (
            WebDriverException,
            StaleElementReferenceException,
        ):
            continue

    ranked.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    return ranked[0][1] if ranked else None


def _attach_resume_to_message(driver, resume_path, timeout=20):
    """Attach the selected resume through LinkedIn's normal message UI.

    First use the HTML file input directly. This avoids opening a native
    Linux file-picker/file-manager window. The visible attachment button is
    only used as a fallback when LinkedIn has not exposed an input.
    """
    resume_path = clean(resume_path)

    if not resume_path or not os.path.isfile(resume_path):
        print(
            f"Attachment failed: resume file does not exist: {resume_path}"
        )
        return False

    filename = os.path.basename(resume_path)
    deadline = time.time() + timeout

    # IMPORTANT: Direct send_keys() to the HTML file input does not open
    # the OS picker. This is the preferred path on Linux.
    file_input = _find_message_file_input(driver)

    if file_input is not None:
        try:
            file_input.send_keys(resume_path)
            print(
                f"Resume selected for attachment: {filename}"
            )
        except (
            WebDriverException,
            StaleElementReferenceException,
        ) as exc:
            print(
                f"Direct file-input selection failed: {exc}"
            )
            file_input = None

    # Fallback only when LinkedIn has not exposed the input yet.
    if file_input is None:
        attachment_button = _find_message_attachment_control(driver)

        if attachment_button is None:
            print(
                "Attachment failed: Message Attachment control "
                "was not exposed."
            )
            return False

        if not _click_visible_action(
            attachment_button,
            "Message Attachment",
        ):
            return False

        time.sleep(0.5)
        file_input = _find_message_file_input(driver)

        while (
            time.time() < deadline
            and file_input is None
        ):
            time.sleep(0.25)
            file_input = _find_message_file_input(driver)

        if file_input is None:
            print(
                "Attachment failed: no HTML file input appeared "
                "after opening attachment UI."
            )
            return False

        try:
            file_input.send_keys(resume_path)
            print(
                f"Resume selected for attachment: {filename}"
            )
        except (
            WebDriverException,
            StaleElementReferenceException,
        ) as exc:
            print(
                f"Attachment failed while selecting resume: {exc}"
            )
            return False

    # Verify that LinkedIn accepted the file before allowing Send.
    verify_deadline = time.time() + timeout

    while time.time() < verify_deadline:
        try:
            value = clean(
                file_input.get_attribute("value")
            )

            if filename.casefold() in value.casefold():
                print(
                    f"Verified resume selected: {filename}"
                )
                return True

        except (
            WebDriverException,
            StaleElementReferenceException,
        ):
            file_input = _find_message_file_input(driver)

        try:
            body = clean(
                driver.find_element(
                    By.TAG_NAME,
                    "body",
                ).text
            ).casefold()

            if filename.casefold() in body:
                print(
                    f"Verified resume attachment in composer UI: "
                    f"{filename}"
                )
                return True

        except WebDriverException:
            pass

        try:
            found = driver.execute_script(
                """
                const wanted = arguments[0].toLowerCase();
                let ok = false;

                function walk(root) {
                    if (!root || !root.querySelectorAll || ok) {
                        return;
                    }

                    for (const e of root.querySelectorAll('*')) {
                        const text =
                            (e.textContent || '') + ' ' +
                            (e.getAttribute('aria-label') || '') + ' ' +
                            (e.getAttribute('title') || '');

                        if (
                            text.toLowerCase().includes(wanted)
                        ) {
                            ok = true;
                            return;
                        }

                        if (e.shadowRoot) {
                            walk(e.shadowRoot);
                        }

                        if (ok) {
                            return;
                        }
                    }
                }

                walk(document);
                return ok;
                """,
                filename,
            )

            if found:
                print(
                    f"Verified resume attachment in rendered "
                    f"composer: {filename}"
                )
                return True

        except WebDriverException:
            pass

        time.sleep(0.35)

    print(
        f"Attachment failed: upload/attachment was not verified "
        f"for {filename}."
    )
    return False


def _find_shadow_dom_elements(driver, selector, visible_only=True):
    """Return matching elements from document and open shadow roots.

    LinkedIn's current message composer is rendered inside the #interop-outlet
    shadow tree in some UI variants. Selenium's normal find_elements() does not
    cross that boundary, so ordinary selectors can report that the composer is
    missing even while it is visibly open.
    """
    try:
        return driver.execute_script("""
            const selector = arguments[0];
            const visibleOnly = arguments[1];
            const results = [];
            const seen = new Set();

            function visible(e) {
                if (!visibleOnly) return true;
                const r = e.getBoundingClientRect();
                const s = getComputedStyle(e);
                return r.width > 0 && r.height > 0 &&
                       s.display !== 'none' && s.visibility !== 'hidden' &&
                       s.opacity !== '0';
            }

            function walk(root) {
                if (!root || !root.querySelectorAll) return;
                let nodes = [];
                try { nodes = root.querySelectorAll(selector); } catch (_) {}
                for (const e of nodes) {
                    if (!visible(e)) continue;
                    if (!seen.has(e)) {
                        seen.add(e);
                        results.push(e);
                    }
                }
                let all = [];
                try { all = root.querySelectorAll('*'); } catch (_) {}
                for (const host of all) {
                    if (host.shadowRoot) walk(host.shadowRoot);
                }
            }

            walk(document);
            return results;
        """, selector, visible_only) or []
    except WebDriverException:
        return []


def _find_message_action_in_shadow(driver, action):
    """Find a visible exact messaging action across light DOM + shadow roots."""
    wanted = clean(action).casefold()
    try:
        return driver.execute_script("""
            const wanted = arguments[0];
            const results = [];
            const seen = new Set();

            function visible(e) {
                const r = e.getBoundingClientRect();
                const s = getComputedStyle(e);
                return r.width > 0 && r.height > 0 &&
                       s.display !== 'none' && s.visibility !== 'hidden' &&
                       s.opacity !== '0' && r.bottom >= 0 && r.top <= window.innerHeight;
            }

            function label(e) {
                return [e.innerText, e.textContent, e.getAttribute('aria-label'), e.getAttribute('title')]
                    .filter(Boolean).join(' ').replace(/\\s+/g, ' ').trim().toLowerCase();
            }

            function walk(root) {
                if (!root || !root.querySelectorAll) return;
                let nodes = [];
                try { nodes = root.querySelectorAll('button, [role="button"], a'); } catch (_) {}
                for (const e of nodes) {
                    if (seen.has(e) || !visible(e)) continue;
                    seen.add(e);
                    const l = label(e);
                    if (l === wanted || l.split(' ').includes(wanted)) results.push(e);
                }
                let all = [];
                try { all = root.querySelectorAll('*'); } catch (_) {}
                for (const host of all) if (host.shadowRoot) walk(host.shadowRoot);
            }
            walk(document);
            return results.length ? results[0] : null;
        """, wanted)
    except WebDriverException:
        return None


def _find_message_composer(driver):
    """Find the actual editable surface in LinkedIn's open message composer.

    LinkedIn's messaging UI has changed markup several times.  In some builds
    the visible editor is a contenteditable node, while the placeholder is
    rendered by a descendant/pseudo-element and the class names differ.
    Selenium's is_displayed() can also be unreliable for dynamically-rendered
    contenteditable nodes.  We therefore use both Selenium selectors and a
    browser-side DOM visibility test, then score candidates by their actual
    screen size and messaging-related attributes.
    """
    def browser_visible(el):
        try:
            return bool(driver.execute_script("""
                const e = arguments[0];
                if (!e) return false;
                const r = e.getBoundingClientRect();
                const s = window.getComputedStyle(e);
                return !!(r.width > 150 && r.height > 20 &&
                    s.display !== 'none' && s.visibility !== 'hidden' &&
                    s.opacity !== '0');
            """, el))
        except (WebDriverException, StaleElementReferenceException):
            return False

    def usable(el):
        try:
            r = el.rect
            size_ok = float(r.get("width", 0) or 0) >= 150 and float(r.get("height", 0) or 0) >= 20
            return size_ok and (el.is_displayed() or browser_visible(el))
        except (WebDriverException, StaleElementReferenceException):
            return False

    def score(el):
        try:
            tag = clean(el.tag_name).casefold()
            role = clean(el.get_attribute("role")).casefold()
            aria = clean(el.get_attribute("aria-label")).casefold()
            placeholder = clean(el.get_attribute("placeholder")).casefold()
            data_placeholder = clean(el.get_attribute("data-placeholder")).casefold()
            cls = clean(el.get_attribute("class")).casefold()
            name = clean(el.get_attribute("name")).casefold()
            ce = clean(el.get_attribute("contenteditable")).casefold()
            combined = f"{role} {aria} {placeholder} {data_placeholder} {cls} {name}"
            r = el.rect
            area = float(r.get("width", 0) or 0) * float(r.get("height", 0) or 0)
            value = 0
            if "msg-form__contenteditable" in cls: value += 5000
            if "msg-form" in cls or "messaging" in combined: value += 1800
            if "write a message" in combined: value += 4000
            if ce in {"true", "plaintext-only"}: value += 3000
            if role == "textbox": value += 2000
            if tag == "textarea": value += 1200
            if area >= 100000: value += 1200
            elif area >= 50000: value += 700
            elif area >= 20000: value += 300
            return value + min(area / 100000.0, 20)
        except (WebDriverException, StaleElementReferenceException):
            return -1

    candidates = []
    seen = set()

    # Current LinkedIn can render the message composer under #interop-outlet's
    # shadow root. Search those nodes explicitly.
    for selector in (
        ".msg-form__contenteditable",
        "[contenteditable='true']",
        "[contenteditable='plaintext-only']",
        "[role='textbox']",
    ):
        for el in _find_shadow_dom_elements(driver, selector, visible_only=True):
            if not usable(el):
                continue
            try:
                key = el.id
            except Exception:
                key = str(id(el))
            if key in seen:
                continue
            seen.add(key)
            candidates.append((score(el) + 8000, el))

    selectors = (
        ".msg-form__contenteditable",
        ".msg-form__contenteditable[contenteditable]",
        "[contenteditable='true']",
        "[contenteditable='plaintext-only']",
        "[role='textbox'][aria-multiline='true']",
        "[role='textbox']",
        "textarea",
    )

    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
        except WebDriverException:
            continue
        for el in elements:
            if not usable(el):
                continue
            try:
                key = el.id
            except Exception:
                key = str(id(el))
            if key in seen:
                continue
            seen.add(key)
            candidates.append((score(el), el))

    # Browser-side query catches dynamically rendered nodes even when Selenium
    # reports their display state incorrectly.
    try:
        js_candidates = driver.execute_script("""
            const selectors = [
                '.msg-form__contenteditable',
                '[contenteditable="true"]',
                '[contenteditable="plaintext-only"]',
                '[role="textbox"]',
                'textarea'
            ];
            const out = [];
            for (const sel of selectors) {
                for (const e of document.querySelectorAll(sel)) {
                    const r = e.getBoundingClientRect();
                    const s = getComputedStyle(e);
                    if (r.width >= 150 && r.height >= 20 &&
                        s.display !== 'none' && s.visibility !== 'hidden' &&
                        s.opacity !== '0') {
                        out.push(e);
                    }
                }
            }
            return out;
        """) or []
        for el in js_candidates:
            if not usable(el):
                continue
            try:
                key = el.id
            except Exception:
                key = str(id(el))
            if key in seen:
                continue
            seen.add(key)
            candidates.append((score(el) + 2500, el))
    except WebDriverException:
        pass

    # The placeholder may be an element inside the editor or an accessible
    # label attached to it. Walk both descendants and contenteditable ancestors.
    xpaths = (
        "//*[@data-placeholder='Write a message...']",
        "//*[@data-placeholder='Write a message…']",
        "//*[@placeholder='Write a message...']",
        "//*[@placeholder='Write a message…']",
        "//*[@aria-label='Write a message...']",
        "//*[@aria-label='Write a message…']",
        "//*[contains(@data-placeholder, 'Write a message')]",
        "//*[contains(@aria-label, 'Write a message')]",
    )
    for xpath in xpaths:
        try:
            markers = driver.find_elements(By.XPATH, xpath)
        except WebDriverException:
            continue
        for marker in markers:
            # Marker itself may be the editor.
            nearby = [marker]
            try:
                nearby += marker.find_elements(
                    By.XPATH,
                    ".//*[self::div or self::p or self::textarea or @role='textbox' or @contenteditable='true' or @contenteditable='plaintext-only']"
                )
            except WebDriverException:
                pass
            try:
                nearby += marker.find_elements(
                    By.XPATH,
                    "./ancestor::*[@contenteditable='true' or @contenteditable='plaintext-only' or @role='textbox'][1]"
                )
            except WebDriverException:
                pass
            for el in nearby:
                if not usable(el):
                    continue
                try:
                    key = el.id
                except Exception:
                    key = str(id(el))
                if key in seen:
                    continue
                seen.add(key)
                candidates.append((score(el) + 4500, el))

    if not candidates:
        # Save DOM diagnostics. This is useful if LinkedIn changes its markup
        # again, without pretending that the composer is absent when it is
        # visibly open.
        try:
            diagnostic = driver.execute_script("""
                return Array.from(document.querySelectorAll('*')).filter(e => {
                    const r = e.getBoundingClientRect();
                    const s = getComputedStyle(e);
                    return r.width >= 150 && r.height >= 20 &&
                           s.display !== 'none' && s.visibility !== 'hidden' &&
                           (e.isContentEditable || e.getAttribute('role') === 'textbox' ||
                            e.getAttribute('data-placeholder') || e.getAttribute('aria-label'));
                }).slice(0, 100).map(e => ({
                    tag:e.tagName,
                    cls:e.className,
                    role:e.getAttribute('role'),
                    ce:e.getAttribute('contenteditable'),
                    placeholder:e.getAttribute('data-placeholder') || e.getAttribute('placeholder'),
                    aria:e.getAttribute('aria-label'),
                    text:(e.innerText || '').slice(0,120),
                    rect:(() => { const r=e.getBoundingClientRect(); return [r.x,r.y,r.width,r.height]; })()
                }));
            """)
            os.makedirs("logs", exist_ok=True)
            with open("logs/message_composer_dom.txt", "w", encoding="utf-8") as fh:
                import json
                fh.write(json.dumps(diagnostic, indent=2, ensure_ascii=False))
        except Exception:
            pass
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]

def _clear_and_type_message(driver, composer, message):
    """Focus the LinkedIn editor, clear it, and enter text using normal UI input."""
    try:
        composer.click()
    except (WebDriverException, StaleElementReferenceException):
        composer = _find_message_composer(driver)
        if composer is None:
            return False
        try:
            composer.click()
        except (WebDriverException, StaleElementReferenceException):
            return False

    # LinkedIn's contenteditable editor often does not implement Selenium's
    # clear() reliably. Ctrl+A + Backspace is the normal keyboard interaction.
    try:
        composer.send_keys(Keys.CONTROL, "a")
        composer.send_keys(Keys.BACKSPACE)
    except (WebDriverException, StaleElementReferenceException):
        try:
            composer = _find_message_composer(driver)
            if composer is None:
                return False
            composer.click()
        except (WebDriverException, StaleElementReferenceException):
            return False

    try:
        # If the first match is a visual placeholder/container, resolve the
        # real editable descendant before typing.
        try:
            if clean(composer.get_attribute("contenteditable")).casefold() not in {"true", "plaintext-only"} and clean(composer.tag_name).casefold() != "textarea":
                descendants = composer.find_elements(By.CSS_SELECTOR, "[contenteditable='true'], [contenteditable='plaintext-only'], textarea, [role='textbox']")
                descendants = [d for d in descendants if d.is_displayed()]
                if descendants:
                    descendants.sort(key=lambda d: (float(d.rect.get("width", 0) or 0) * float(d.rect.get("height", 0) or 0)), reverse=True)
                    composer = descendants[0]
        except (WebDriverException, StaleElementReferenceException):
            pass

        composer.click()
        composer.send_keys(message)
    except (WebDriverException, StaleElementReferenceException):
        try:
            composer = _find_message_composer(driver)
            if composer is None:
                return False
            composer.click()
            composer.send_keys(message)
        except (WebDriverException, StaleElementReferenceException):
            return False

    # Verify that some of the typed message actually reached the editor.
    try:
        value = clean(composer.get_attribute("value"))
        text = clean(composer.text)
        inner_text = clean(composer.get_attribute("innerText"))
        current = f"{value} {text} {inner_text}".casefold()
        if clean(message).casefold()[:40] in current or len(current) >= min(40, len(clean(message))):
            return True
    except (WebDriverException, StaleElementReferenceException):
        # A DOM rebuild after typing is normal; the absence of a readable value
        # does not by itself mean typing failed.
        return True

    return False

def _find_message_send_control(driver):
    """Find the actual Send control in the open LinkedIn message composer.

    LinkedIn may render Send as:
      * text: "Send"
      * icon-only button with aria-label/title "Send"
      * a control whose data-control-name/class contains "send"
      * a button inside a shadow root

    This helper is intentionally restricted to visible/enabled controls whose
    accessible/structural metadata identifies them as a Send control. It does
    not click arbitrary icons.
    """
    # First use normal Selenium DOM controls.
    try:
        elements = driver.find_elements(
            By.CSS_SELECTOR,
            "button, [role='button'], a, input[type='button'], input[type='submit']",
        )
    except WebDriverException:
        elements = []

    def is_good(element):
        try:
            if not element.is_displayed() or not element.is_enabled():
                return False

            text = clean(element.text).casefold()
            aria = clean(element.get_attribute("aria-label")).casefold()
            title = clean(element.get_attribute("title")).casefold()
            name = clean(element.get_attribute("name")).casefold()
            data_name = clean(element.get_attribute("data-control-name")).casefold()
            data_test = clean(element.get_attribute("data-testid")).casefold()
            cls = clean(element.get_attribute("class")).casefold()

            # Exact accessible labels are strongest, including icon-only buttons.
            if text == "send" or aria == "send" or title == "send":
                return True

            # Common LinkedIn structural identifiers for icon-only Send.
            structural = " ".join(
                (name, data_name, data_test, cls)
            )
            if re.search(r"(^|[-_\\s])send([-_\\s]|$)", structural):
                return True

            if "send-button" in structural or "msg-form__send" in structural:
                return True

            return False
        except (WebDriverException, StaleElementReferenceException):
            return False

    # Prefer controls physically near the visible composer.
    composer = _find_message_composer(driver)
    composer_rect = None
    if composer is not None:
        try:
            composer_rect = composer.rect
        except (WebDriverException, StaleElementReferenceException):
            composer_rect = None

    candidates = [e for e in elements if is_good(e)]
    if candidates:
        if composer_rect:
            cx = float(composer_rect.get("x", 0) or 0) + float(composer_rect.get("width", 0) or 0) / 2
            cy = float(composer_rect.get("y", 0) or 0) + float(composer_rect.get("height", 0) or 0) / 2

            def distance(e):
                try:
                    r = e.rect
                    ex = float(r.get("x", 0) or 0) + float(r.get("width", 0) or 0) / 2
                    ey = float(r.get("y", 0) or 0) + float(r.get("height", 0) or 0) / 2
                    return abs(ex - cx) + abs(ey - cy)
                except Exception:
                    return 10**9

            candidates.sort(key=distance)
        return candidates[0]

    # Finally inspect shadow DOM. The JS returns a real WebElement.
    try:
        return driver.execute_script(r"""
            const results = [];
            const seen = new Set();

            function visible(e) {
                const r = e.getBoundingClientRect();
                const s = getComputedStyle(e);
                return r.width > 0 && r.height > 0 &&
                       s.display !== 'none' &&
                       s.visibility !== 'hidden' &&
                       s.opacity !== '0' &&
                       r.bottom >= 0 && r.top <= window.innerHeight;
            }

            function enabled(e) {
                return !e.disabled && e.getAttribute('aria-disabled') !== 'true';
            }

            function meta(e) {
                return [
                    e.innerText,
                    e.textContent,
                    e.getAttribute('aria-label'),
                    e.getAttribute('title'),
                    e.getAttribute('name'),
                    e.getAttribute('data-control-name'),
                    e.getAttribute('data-testid'),
                    e.className
                ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
            }

            function isSend(e) {
                const m = meta(e);
                if (m === 'send') return true;
                if (/\bsend\b/.test(m) &&
                    (m.includes('button') || m.includes('msg-form') ||
                     m.includes('send-button') || m.includes('control'))) {
                    return true;
                }
                return false;
            }

            function walk(root) {
                if (!root || !root.querySelectorAll) return;
                let nodes = [];
                try {
                    nodes = root.querySelectorAll(
                        'button, [role="button"], a, input[type="button"], input[type="submit"]'
                    );
                } catch (_) {}

                for (const e of nodes) {
                    if (seen.has(e) || !visible(e) || !enabled(e)) continue;
                    seen.add(e);
                    if (isSend(e)) results.push(e);
                }

                let all = [];
                try { all = root.querySelectorAll('*'); } catch (_) {}
                for (const host of all) {
                    if (host.shadowRoot) walk(host.shadowRoot);
                }
            }

            walk(document);
            return results.length ? results[0] : null;
        """)
    except WebDriverException:
        return None


def send_linkedin_message(
    driver,
    profile_url,
    message,
    attachment_path=None,
    require_attachment=False,
):
    """Send a LinkedIn message through the normal visible UI.

    Accepted-connection referral flow:
      1. Open profile.
      2. Click the profile's Message action.
      3. Wait for the actual message editor to render.
      4. Type the referral message.
      5. Attach the selected resume when requested.
      6. Wait for an enabled Send control and click it.

    When ``require_attachment`` is True, the message is never sent unless the
    resume attachment has been successfully selected/verified.
    """
    if not profile_url or not message:
        return False

    if require_attachment and not attachment_path:
        print("Message send blocked: required attachment path is missing.")
        return False

    try:
        driver.get(profile_url)
        time.sleep(PEOPLE_WAIT_SECONDS)

        if not _click_profile_action(driver, "Message", timeout=12):
            print("Message button was not exposed/clickable in the normal profile UI.")
            return False

        # The Message popup is asynchronous. Do not assume a fixed sleep is
        # enough; poll the actual DOM for the large contenteditable/textarea.
        composer = None
        composer_deadline = time.time() + 20
        while time.time() < composer_deadline:
            composer = _find_message_composer(driver)
            if composer is not None:
                break
            time.sleep(0.35)

        if composer is None:
            print("Message composer was not exposed after waiting for LinkedIn messaging UI.")
            try:
                driver.save_screenshot("logs/message_composer_not_exposed.png")
            except WebDriverException:
                pass
            return False

        print("Message composer exposed; entering referral message.")
        if not _clear_and_type_message(driver, composer, message):
            print("Message composer was found, but referral text could not be entered.")
            try:
                driver.save_screenshot("logs/message_text_entry_failed.png")
            except WebDriverException:
                pass
            return False

        time.sleep(0.8)

        if attachment_path:
            attached = _attach_resume_to_message(
                driver,
                attachment_path,
                timeout=20,
            )
            if not attached:
                print(
                    "Message send blocked: required resume attachment "
                    "could not be verified."
                )
                try:
                    driver.save_screenshot("logs/message_attachment_failed.png")
                except WebDriverException:
                    pass
                return False

        # LinkedIn can render Send as text OR as an icon-only button.
        send_button = None
        send_deadline = time.time() + 15
        while time.time() < send_deadline:
            send_button = _find_message_send_control(driver)
            if send_button is not None:
                try:
                    if send_button.is_enabled():
                        break
                except (WebDriverException, StaleElementReferenceException):
                    # A shadow-root element can be rebuilt between polls.
                    pass
            send_button = None
            time.sleep(0.35)

        if send_button is None:
            print("Message composer Send control was not exposed/enabled after waiting.")
            try:
                driver.save_screenshot("logs/message_send_not_exposed.png")
            except WebDriverException:
                pass
            return False

        print("FOUND MESSAGE SEND CONTROL (text or icon)")
        if not _click_visible_action(send_button, "Message Send"):
            print("Could not click the LinkedIn Message Send button.")
            return False

        time.sleep(2)
        print(
            "Message send click completed"
            + (" with resume attachment." if attachment_path else ".")
        )
        return True

    except (WebDriverException, StaleElementReferenceException) as exc:
        print(f"LinkedIn message UI failed: {exc}")
        return False
    except Exception as exc:
        print(f"LinkedIn message failed: {exc}")
        return False

def _find_visible_action(driver, phrases=(), exact=()):
    """Find a visible/enabled normal LinkedIn UI action by text/aria/title."""
    try:
        elements = driver.find_elements(
            By.CSS_SELECTOR,
            "button, [role='button'], [role='menuitem'], a",
        )
    except WebDriverException:
        return None

    exact_set = {clean(x).lower() for x in exact}
    phrase_list = [clean(x).lower() for x in phrases]

    for element in elements:
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue

            text = clean(element.text)
            aria = clean(element.get_attribute("aria-label"))
            title = clean(element.get_attribute("title"))
            low_text = text.lower()
            combined = f"{text} {aria} {title}".lower()

            if low_text in exact_set:
                return element

            if any(phrase in combined for phrase in phrase_list):
                return element

        except (WebDriverException, StaleElementReferenceException):
            continue

    return None


def _click_visible_action(element, label):
    """Click a visible LinkedIn action using normal browser interaction."""
    if element is None:
        print(f"Clicking {label} failed: element is None")
        return False

    try:
        print(f"Clicking {label}...")
        element.click()
        print(f"Clicked {label} with Selenium.")
        return True
    except (WebDriverException, StaleElementReferenceException) as exc:
        print(f"Selenium click failed for {label}: {exc}")

    # Retry through the normal user interaction chain. This is still the
    # same visible rendered element; it does not bypass LinkedIn checks.
    try:
        driver = element.parent
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center', inline:'center'});",
            element,
        )
    except (WebDriverException, StaleElementReferenceException):
        pass

    try:
        ActionChains(element.parent).move_to_element(element).pause(0.2).click().perform()
        print(f"Clicked {label} with ActionChains.")
        return True
    except (WebDriverException, StaleElementReferenceException) as exc:
        print(f"ActionChains click failed for {label}: {exc}")
        return False


def _visible_button_by_exact_text(driver, label):
    """
    Find the exact visible LinkedIn control for `label`.

    LinkedIn may render the profile action as a <button>, <a>, or an
    element with role="button".  Do not assume one tag.  The returned
    element must be visible, enabled, and have the exact rendered label
    (or exact aria/title label).
    """
    wanted = clean(label).casefold()

    # Include bare <a>: LinkedIn profile actions are not always rendered as
    # <button> elements. This was the reason the Message control could be
    # detected by the page-state check but not located for the click.
    selectors = (
        "button",
        "a",
        "[role='button']",
    )

    elements = []
    seen = set()
    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                try:
                    key = element.id
                except Exception:
                    key = id(element)
                if key not in seen:
                    seen.add(key)
                    elements.append(element)
        except WebDriverException:
            continue

    # Pass 1: exact rendered text / aria / title.
    for element in elements:
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue

            text = clean(element.text).casefold()
            aria = clean(element.get_attribute("aria-label")).casefold()
            title = clean(element.get_attribute("title")).casefold()

            if text == wanted or aria == wanted or title == wanted:
                return element
        except (WebDriverException, StaleElementReferenceException):
            continue

    # Pass 2: LinkedIn often puts the label in a nested span. Return the
    # nearest clickable element whose full rendered text is exactly the label.
    for element in elements:
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue

            text = clean(element.text).casefold()
            aria = clean(element.get_attribute("aria-label")).casefold()
            title = clean(element.get_attribute("title")).casefold()

            if wanted == text or wanted == aria or wanted == title:
                return element
        except (WebDriverException, StaleElementReferenceException):
            continue

    # Pass 3: direct XPath text/attribute lookup. This catches controls where
    # Selenium's element.text does not expose the nested visible span exactly.
    xpath = (
        "//*[self::button or self::a or @role='button']["
        "normalize-space(.)=" + repr(wanted) + " or "
        "translate(normalize-space(@aria-label),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')=" + repr(wanted) + " or "
        "translate(normalize-space(@title),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')=" + repr(wanted) + "]"
    )

    # Python repr() is not XPath string escaping. Use a safe literal builder
    # for the normal labels used here.
    def xpath_literal(value):
        if "'" not in value:
            return "'" + value + "'"
        if '"' not in value:
            return '"' + value + '"'
        parts = value.split("'")
        return "concat(" + ", \"'\", ".join("'" + p + "'" for p in parts) + ")"

    lit = xpath_literal(wanted)
    xpath = (
        "//*[self::button or self::a or @role='button']["
        f"translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')={lit} "
        "or "
        f"translate(normalize-space(@aria-label),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')={lit} "
        "or "
        f"translate(normalize-space(@title),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')={lit}"
        "]"
    )

    try:
        xpath_elements = driver.find_elements(By.XPATH, xpath)
    except WebDriverException:
        xpath_elements = []

    for element in xpath_elements:
        try:
            if element.is_displayed() and element.is_enabled():
                return element
        except (WebDriverException, StaleElementReferenceException):
            continue

    return None


def _click_profile_action(driver, label, timeout=12):
    """Repeatedly locate and click one exact visible profile action."""
    deadline = time.time() + timeout
    last_error = ""

    while time.time() < deadline:
        element = _visible_button_by_exact_text(driver, label)
        if element is None:
            time.sleep(0.25)
            continue

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                element,
            )
        except WebDriverException:
            pass

        # Selenium's documented clickable condition checks visible + enabled.
        # Re-locate the element after the SPA has settled to avoid stale refs.
        try:
            if element.is_displayed() and element.is_enabled():
                print(f"FOUND EXACT PROFILE ACTION: {label}")
                print(f"CLICKING PROFILE ACTION: {label}")
                element.click()
                print(f"Clicked profile action with Selenium: {label}")
                return True
        except (WebDriverException, StaleElementReferenceException) as exc:
            last_error = str(exc)
            print(f"Selenium click failed for profile action '{label}': {exc}")

        # Normal mouse-style retry, still targeting the same rendered control.
        try:
            fresh = _visible_button_by_exact_text(driver, label)
            if fresh is not None:
                ActionChains(driver).move_to_element(fresh).pause(0.2).click().perform()
                print(f"Clicked profile action with ActionChains: {label}")
                return True
        except (WebDriverException, StaleElementReferenceException) as exc:
            last_error = str(exc)
            print(f"ActionChains click failed for profile action '{label}': {exc}")

        time.sleep(0.35)

    print(f"FAILED TO CLICK PROFILE ACTION '{label}'. {last_error}")
    try:
        driver.save_screenshot(f"logs/profile_action_{clean(label).lower().replace(' ', '_')}_failed.png")
    except WebDriverException:
        pass
    return False

def _click_invitation_button(driver, label, timeout=15):
    """
    Click the exact invitation-popup button and verify that the click caused
    the button to disappear/change.

    This is intentionally separate from the generic page-action click because
    the invitation popup has two visually similar choices and the normal
    generic matcher was too easy to confuse.
    """
    deadline = time.time() + timeout
    last_error = ""

    while time.time() < deadline:
        button = _visible_button_by_exact_text(driver, label)

        if button is None:
            time.sleep(0.25)
            continue

        try:
            print(
                f"FOUND INVITATION BUTTON: text={clean(button.text)!r} "
                f"| aria={clean(button.get_attribute('aria-label'))!r}"
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                button,
            )
        except WebDriverException:
            pass

        # First attempt: Selenium's real element click.
        try:
            print(f"CLICKING INVITATION BUTTON: {label}")
            button.click()
            print(f"Selenium clicked invitation button: {label}")
        except (WebDriverException, StaleElementReferenceException) as exc:
            last_error = str(exc)
            print(f"Selenium click failed for '{label}': {exc}")

            # Second attempt: normal mouse-style interaction.
            try:
                ActionChains(driver).move_to_element(button).pause(0.2).click().perform()
                print(f"ActionChains clicked invitation button: {label}")
            except (WebDriverException, StaleElementReferenceException) as exc2:
                last_error = str(exc2)
                print(f"ActionChains click failed for '{label}': {exc2}")
                time.sleep(0.4)
                continue

        # The choice should disappear once LinkedIn accepts the click.
        verify_deadline = time.time() + 3
        while time.time() < verify_deadline:
            if _visible_button_by_exact_text(driver, label) is None:
                print(f"VERIFIED: invitation button '{label}' disappeared after click.")
                return True
            time.sleep(0.2)

        # If it is still visible, do not pretend the click worked.
        print(
            f"Invitation button '{label}' is STILL visible after click. "
            "Retrying the exact same button."
        )
        time.sleep(0.4)

    print(
        f"FAILED TO CLICK INVITATION BUTTON '{label}' within {timeout}s. "
        f"{last_error}"
    )
    try:
        driver.save_screenshot("logs/invitation_button_click_failed.png")
        print("DEBUG screenshot saved: logs/invitation_button_click_failed.png")
    except WebDriverException:
        pass

    return False


def _click_exact_button(driver, label, timeout=12):
    """Click an exact visible LinkedIn button, with verification."""
    return _click_invitation_button(driver, label, timeout=timeout)


def _button_exists(driver, label):
    return _visible_button_by_exact_text(driver, label) is not None


def _find_invitation_choice_dom(driver, label):
    """Find an exact invitation choice in LinkedIn's rendered UI.

    The post-Connect sheet is a SPA overlay.  LinkedIn can render the visible
    label on a nested span/div, and some UI variants can place controls behind
    an open shadow root.  We therefore search normal DOM, open shadow roots,
    and then walk upward to the nearest genuinely clickable ancestor.
    """
    wanted = clean(label).casefold()
    if not wanted:
        return None

    # 1) Normal Selenium controls.
    element = _visible_button_by_exact_text(driver, label)
    if element is not None:
        return element

    # 2) Exact text nodes/elements in the normal DOM.  Prefer the smallest
    # visible node whose own text is exactly the requested label.
    try:
        xpath_variants = (
            f"//*[normalize-space(.)={_xpath_literal(label)}]",
            f"//*[@aria-label={_xpath_literal(label)}]",
            f"//*[@title={_xpath_literal(label)}]",
        )
        candidates = []
        for xpath in xpath_variants:
            candidates.extend(driver.find_elements(By.XPATH, xpath))

        for candidate in candidates:
            try:
                if not candidate.is_displayed() or not candidate.is_enabled():
                    continue
                text = clean(candidate.text).casefold()
                aria = clean(candidate.get_attribute('aria-label')).casefold()
                title = clean(candidate.get_attribute('title')).casefold()
                if text == wanted or aria == wanted or title == wanted:
                    return _nearest_clickable_ancestor(driver, candidate, label)
            except (WebDriverException, StaleElementReferenceException):
                continue
    except WebDriverException:
        pass

    # 3) Browser-side recursive DOM search, including OPEN shadow roots.
    # Return the actual DOM element so Selenium can perform a normal click.
    try:
        script = r"""
        const wanted = String(arguments[0] || '').trim().toLowerCase();

        function norm(v) {
            return String(v || '').replace(/\s+/g, ' ').trim().toLowerCase();
        }

        function visible(el) {
            if (!el || !el.getBoundingClientRect) return false;
            const r = el.getBoundingClientRect();
            const st = getComputedStyle(el);
            return r.width > 0 && r.height > 0 &&
                   st.display !== 'none' &&
                   st.visibility !== 'hidden' &&
                   st.opacity !== '0';
        }

        function exact(el) {
            if (!el || !visible(el)) return false;
            const text = norm(el.innerText || el.textContent);
            const aria = norm(el.getAttribute && el.getAttribute('aria-label'));
            const title = norm(el.getAttribute && el.getAttribute('title'));
            return text === wanted || aria === wanted || title === wanted;
        }

        function clickable(el) {
            if (!el || !visible(el)) return false;
            const tag = (el.tagName || '').toLowerCase();
            const role = norm(el.getAttribute && el.getAttribute('role'));
            const type = norm(el.getAttribute && el.getAttribute('type'));
            const aria = norm(el.getAttribute && el.getAttribute('aria-label'));
            const tab = el.getAttribute && el.getAttribute('tabindex');
            const dataControl = el.getAttribute && (
                el.getAttribute('data-control-name') ||
                el.getAttribute('data-test-id') ||
                el.getAttribute('data-view-name') ||
                el.getAttribute('data-is-focusable')
            );
            const cursor = getComputedStyle(el).cursor;

            return tag === 'button' || tag === 'a' ||
                   role === 'button' || type === 'button' ||
                   aria === wanted || tab !== null ||
                   !!dataControl || cursor === 'pointer';
        }

        function nearest(el) {
            let cur = el;
            for (let i = 0; i < 12 && cur; i++, cur = cur.parentElement) {
                if (clickable(cur)) return cur;
            }
            return null;
        }

        function walk(root) {
            const nodes = root.querySelectorAll ?
                Array.from(root.querySelectorAll('*')) : [];

            // Prefer exact leaf-ish nodes first. This prevents the whole modal
            // container from being selected because it contains the same text.
            for (const node of nodes) {
                if (exact(node)) {
                    const c = nearest(node);
                    if (c) return c;
                }
            }

            // Search open shadow roots recursively.
            for (const node of nodes) {
                if (node.shadowRoot) {
                    const found = walk(node.shadowRoot);
                    if (found) return found;
                }
            }
            return null;
        }

        return walk(document);
        """
        result = driver.execute_script(script, label)
        if result is not None:
            return result
    except WebDriverException:
        pass

    return None


def _nearest_clickable_ancestor(driver, element, label):
    """Return a clickable ancestor for an exact-text child element."""
    try:
        script = r"""
        const start = arguments[0];
        function visible(el) {
            if (!el || !el.getBoundingClientRect) return false;
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 &&
                   s.display !== 'none' && s.visibility !== 'hidden';
        }
        function clickable(el) {
            if (!visible(el)) return false;
            const tag = (el.tagName || '').toLowerCase();
            const role = (el.getAttribute('role') || '').toLowerCase();
            const type = (el.getAttribute('type') || '').toLowerCase();
            const cursor = getComputedStyle(el).cursor;
            return tag === 'button' || tag === 'a' || role === 'button' ||
                   type === 'button' || el.hasAttribute('tabindex') ||
                   cursor === 'pointer';
        }
        let cur = start;
        for (let i = 0; i < 12 && cur; i++, cur = cur.parentElement) {
            if (clickable(cur)) return cur;
        }
        return start;
        """
        return driver.execute_script(script, element) or element
    except WebDriverException:
        return element

def _xpath_literal(value):
    """Safely quote a string for an XPath literal."""
    value = str(value)
    if "'" not in value:
        return "'" + value + "'"
    if '"' not in value:
        return '"' + value + '"'
    parts = value.split("'")
    return "concat(" + ", \"'\", ".join("'" + p + "'" for p in parts) + ")"


def _click_invitation_choice(driver, label, timeout=20):
    """Locate and click the exact invitation-sheet choice.

    This is deliberately broader than the profile-action finder because the
    post-Connect sheet is dynamically inserted by LinkedIn and its clickable
    wrapper is not guaranteed to be a <button> element.
    """
    deadline = time.time() + timeout
    last_error = ''

    while time.time() < deadline:
        button = _find_invitation_choice_dom(driver, label)
        if button is None:
            time.sleep(0.25)
            continue

        try:
            text = clean(button.text)
            aria = clean(button.get_attribute('aria-label'))
            print(
                f"FOUND INVITATION CHOICE: requested={label!r} "
                f"| text={text!r} | aria={aria!r} | tag={button.tag_name!r}"
            )
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                button,
            )
        except WebDriverException:
            pass

        # Normal Selenium click first.
        try:
            print(f"CLICKING INVITATION CHOICE: {label}")
            button.click()
            print(f"Selenium clicked invitation choice: {label}")
        except (WebDriverException, StaleElementReferenceException) as exc:
            last_error = str(exc)
            print(f"Selenium invitation click failed for '{label}': {exc}")
            try:
                fresh = _find_invitation_choice_dom(driver, label)
                if fresh is not None:
                    ActionChains(driver).move_to_element(fresh).pause(0.2).click().perform()
                    print(f"ActionChains clicked invitation choice: {label}")
                else:
                    time.sleep(0.4)
                    continue
            except (WebDriverException, StaleElementReferenceException) as exc2:
                last_error = str(exc2)
                print(f"ActionChains invitation click failed for '{label}': {exc2}")
                time.sleep(0.4)
                continue

        # Wait for the exact choice to disappear. This is the real success
        # signal, not merely the fact that click() returned.
        verify_deadline = time.time() + 5
        while time.time() < verify_deadline:
            if _find_invitation_choice_dom(driver, label) is None:
                print(f"VERIFIED: invitation choice '{label}' disappeared after click.")
                return True
            time.sleep(0.25)

        print(
            f"Invitation choice '{label}' is still visible after click; "
            "re-locating and retrying."
        )
        time.sleep(0.4)

    print(f"FAILED TO CLICK INVITATION CHOICE '{label}' within {timeout}s. {last_error}")
    try:
        driver.save_screenshot('logs/invitation_choice_failed.png')
        print('DEBUG screenshot saved: logs/invitation_choice_failed.png')
    except WebDriverException:
        pass
    return False


def _invitation_sheet_visible(driver):
    """Return whether LinkedIn's post-Connect invitation sheet is visible."""
    for label in ('Send without a note', 'Add a note'):
        if _find_invitation_choice_dom(driver, label) is not None:
            return True
    return False


def _wait_for_invitation_step(driver, timeout=20):
    """Wait for LinkedIn's post-Connect invitation sheet.

    LinkedIn can insert this sheet asynchronously after Connect is clicked.
    We therefore inspect the rendered DOM repeatedly instead of relying only
    on a fixed sleep or a specific HTML tag.
    """
    deadline = time.time() + timeout

    while time.time() < deadline:
        without_note = _find_invitation_choice_dom(driver, 'Send without a note')
        if without_note is not None:
            print("Detected exact 'Send without a note' invitation choice.")
            return 'send_without_note', without_note

        add_note = _find_invitation_choice_dom(driver, 'Add a note')
        if add_note is not None:
            print("Detected exact 'Add a note' invitation choice.")
            return 'add_note', add_note

        # Diagnostic: detect the popup by its heading/body text even if the
        # choices have not become clickable yet.
        try:
            body = clean(driver.find_element(By.TAG_NAME, 'body').text)
            low = body.casefold()
            if 'add a note to your invitation' in low or 'send without a note' in low:
                print('Invitation popup text detected; waiting for its clickable choices...')
        except WebDriverException:
            pass

        time.sleep(0.25)

    try:
        driver.save_screenshot('logs/connection_dialog_timeout.png')
        print('DEBUG screenshot saved: logs/connection_dialog_timeout.png')
        body = clean(driver.find_element(By.TAG_NAME, 'body').text)
        print('DEBUG visible body tail:', body[-1200:])
    except WebDriverException:
        pass

    return None, None

def _find_connection_note_field(driver):
    """Find the visible note composer in LinkedIn's Add a note dialog."""
    try:
        fields = driver.find_elements(
            By.CSS_SELECTOR,
            "textarea, [contenteditable='true'], input[type='text']",
        )
    except WebDriverException:
        return None

    for field in fields:
        try:
            if not field.is_displayed() or not field.is_enabled():
                continue

            aria = clean(field.get_attribute("aria-label")).lower()
            placeholder = clean(
                field.get_attribute("placeholder")
            ).lower()
            name = clean(field.get_attribute("name")).lower()
            combined = f"{aria} {placeholder} {name}"

            # A textarea is normally the LinkedIn note field. For other
            # inputs, require note-related metadata to avoid filling search.
            if field.tag_name.lower() != "textarea":
                if "note" not in combined:
                    continue

            return field

        except (WebDriverException, StaleElementReferenceException):
            continue

    return None


def _enter_connection_note(driver, note):
    """Enter the supplied connection note into the normal note composer."""
    note = clean(note)[:200]
    if not note:
        return True

    deadline = time.time() + 8

    while time.time() < deadline:
        field = _find_connection_note_field(driver)

        if field is not None:
            try:
                field.click()
                try:
                    field.clear()
                except WebDriverException:
                    pass

                field.send_keys(note)
                print("Connection note entered.")
                return True

            except (WebDriverException, StaleElementReferenceException):
                pass

        time.sleep(0.4)

    return False


def _find_final_invitation_send(driver, timeout=12):
    """Find the final invitation Send button after the intermediate choice."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        # Prefer the exact final labels used by LinkedIn.
        for label in ("Send invitation", "Send connection request"):
            button = _visible_button_by_exact_text(driver, label)
            if button is not None:
                return button

        # A plain Send is valid only while the invitation confirmation is
        # visible. Avoid using arbitrary page Send buttons when no invitation
        # controls are present.
        if _button_exists(driver, "Send"):
            try:
                body = clean(driver.find_element(By.TAG_NAME, "body").text).lower()
            except WebDriverException:
                body = ""
            invitation_words = (
                "send without a note",
                "add a note",
                "invitation",
                "connection request",
            )
            if any(word in body for word in invitation_words):
                return _visible_button_by_exact_text(driver, "Send")

        time.sleep(0.25)

    return None

def _profile_has_message_and_connect(driver):
    """Return True when both Message and Connect are simultaneously visible."""
    has_message = False
    has_connect = False

    try:
        elements = driver.find_elements(
            By.CSS_SELECTOR,
            "button, [role='button'], a",
        )
    except WebDriverException:
        return False

    for element in elements:
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue

            text = clean(element.text).casefold()
            aria = clean(element.get_attribute("aria-label")).casefold()
            title = clean(element.get_attribute("title")).casefold()

            if text == "message" or aria == "message" or title == "message":
                has_message = True

            if (
                text == "connect"
                or ("invite " in aria and "to connect" in aria)
                or title == "connect"
            ):
                has_connect = True

            if has_message and has_connect:
                return True

        except (WebDriverException, StaleElementReferenceException):
            continue

    return False


def _message_requires_premium(driver):
    """Return True when LinkedIn shows a Premium/InMail gate for the profile message action."""
    try:
        body = clean(driver.find_element(By.TAG_NAME, "body").text).lower()
    except WebDriverException:
        return False

    premium_markers = (
        "premium",
        "inmail",
        "upgrade to premium",
        "try premium",
        "get premium",
        "premium subscription",
    )
    return any(marker in body for marker in premium_markers)


def recover_referral_send_lock(job_id, person):
    """Remove a referral pre-send lock only by explicit operator request.

    This is intentionally NOT automatic. A ``started`` lock can mean the
    message was typed and the browser was interrupted before Excel recorded it.
    Automatic retry could therefore create a duplicate referral. Use this
    explicit recovery command only after visually confirming that LinkedIn did
    not send the referral.

    If an exact Referral Request/Sent record exists in Messages, refuse to
    unlock so the duplicate protection remains intact.
    """
    job_id = clean(job_id)
    person = clean(person)

    if not job_id or not person:
        print("Recovery blocked: Job ID and Person are required.")
        return False

    if _referral_message_already_sent(job_id, person):
        print(
            f"Recovery blocked: a Referral Request/Sent record already exists "
            f"for {person} | job={job_id}."
        )
        return False

    data = _load_referral_send_locks()
    if data is None:
        print("Recovery blocked: referral safety state could not be read.")
        return False

    key = _referral_lock_key(job_id, person)
    entry = data.get(key)

    if not entry:
        print(f"No referral lock exists for {person} | job={job_id}.")
        return True

    print(
        f"Removing explicit recovery lock for {person} | job={job_id} "
        f"(state={entry.get('state', 'unknown')})."
    )

    del data[key]

    try:
        _save_referral_send_locks(data)
    except Exception as exc:
        print(f"Recovery failed: could not save lock file: {exc}")
        return False

    print(f"Referral lock removed for {person} | job={job_id}.")
    return True


def send_connection_request(driver, profile_url, note=None, message_text=None):
    """Send exactly one LinkedIn connection invitation through normal UI.

    IMPORTANT: message_text is intentionally ignored for connection outreach.
    A referral request is a separate workflow and is sent only after the live
    profile is later confirmed as connected. This prevents a profile showing
    both Message and Connect from receiving the referral prematurely.
    """
    if not profile_url:
        return False

    note = clean(note)[:200]

    try:
        print(f"Opening profile for connection request: {profile_url}")
        driver.get(profile_url)
        time.sleep(PEOPLE_WAIT_SECONDS)

        connect_button = _find_visible_action(
            driver,
            phrases=("invite ", "to connect"),
            exact=("connect",),
        )

        if connect_button is not None:
            text = clean(connect_button.text).lower()
            aria = clean(connect_button.get_attribute("aria-label")).lower()
            combined = f"{text} {aria}"
            if not (
                text == "connect"
                or ("invite " in aria and "to connect" in aria)
            ):
                connect_button = None
            if any(x in combined for x in ("connected", "pending", "invitation sent")):
                connect_button = None

        if connect_button is None:
            print("RESULT: Connect action was not exposed by the normal LinkedIn profile UI.")
            try:
                driver.save_screenshot("logs/connection_connect_debug.png")
            except WebDriverException:
                pass
            return False

        print(
            f"FOUND CONNECT: text={clean(connect_button.text)!r} "
            f"| aria={clean(connect_button.get_attribute('aria-label'))!r}"
        )

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                connect_button,
            )
        except WebDriverException:
            pass

        if not _click_visible_action(connect_button, "Connect"):
            return False

        # ---------------------------------------------------------
        # LinkedIn invitation choice.
        # ---------------------------------------------------------
        step, action = _wait_for_invitation_step(driver, timeout=15)

        if action is None:
            print("RESULT: Connect clicked, but no invitation action appeared.")
            try:
                driver.save_screenshot("logs/connection_dialog_debug.png")
            except WebDriverException:
                pass
            return False

        # High-value contact: use Add a note only when the caller supplied
        # a real note. Ordinary contacts deliberately use no note.
        if note and step in {"add_note", "send_without_note"}:
            if step != "add_note":
                # The UI exposed the no-note action first. For a high-value
                # contact we wait briefly for Add a note rather than sending
                # without the note.
                add_note = _visible_button_by_exact_text(driver, "Add a note")
                if add_note is None:
                    print("High-value contact: Add a note was not exposed; not sending without note.")
                    return False
                action = add_note

            if not _click_invitation_choice(driver, "Add a note", timeout=20):
                return False

            if not _enter_connection_note(driver, note):
                print("RESULT: Add a note dialog opened, but note field was not usable.")
                return False

            # After entering a note LinkedIn normally exposes Send.
            if not _click_invitation_choice(driver, "Send", timeout=12):
                return False

        elif not note and step == "send_without_note":
            # THIS is the critical normal path:
            # Connect -> exact visible 'Send without a note' button.
            print("Normal contact: sending WITHOUT a note.")
            if not _click_invitation_choice(driver, "Send without a note", timeout=20):
                return False

        elif not note and step == "add_note":
            # LinkedIn exposed the two-choice dialog but our exact finder
            # returned Add a note first. Never click Add a note for a normal
            # contact; explicitly find the no-note button and click it.
            print("Normal contact: Add a note is visible; looking specifically for Send without a note.")
            if not _click_invitation_choice(driver, "Send without a note", timeout=20):
                return False

        elif step == "send":
            if not _click_visible_action(action, "Send"):
                return False

        else:
            print(f"RESULT: Unsupported invitation state: {step}")
            return False

        # ---------------------------------------------------------
        # Some LinkedIn variants expose a second final confirmation.
        # Only inspect/click a plain "Send" while the invitation sheet itself
        # is still visible. Never click an arbitrary Send control on the profile.
        # ---------------------------------------------------------
        print("Checking for final invitation confirmation...")
        if _invitation_sheet_visible(driver):
            final_send = _find_final_invitation_send(driver, timeout=5)
            if final_send is not None:
                print("Final invitation Send detected. Clicking it now...")
                if not _click_visible_action(final_send, "final Send"):
                    return False
            else:
                print(
                    "Invitation sheet is still visible but no final Send "
                    "control was exposed; treating connection request as failed."
                )
                return False
        else:
            print("Invitation sheet closed; continuing to verification.")

        time.sleep(2)

        state = get_profile_connection_state(driver, profile_url)

        if state == "pending":
            print("SUCCESS: Connection request is pending.")
            print(
                "Connection request was sent "
                + ("with a note." if note else "without a note.")
            )
            return True

        try:
            body_text = clean(driver.find_element(By.TAG_NAME, "body").text).lower()
        except WebDriverException:
            body_text = ""

        if any(
            text in body_text
            for text in (
                "invitation sent",
                "request sent",
                "invitation pending",
            )
        ):
            print("SUCCESS: LinkedIn indicates the connection request was sent.")
            return True

        print(f"RESULT: Send flow completed, but LinkedIn returned profile state: {state}")
        try:
            driver.save_screenshot("logs/connection_verify_debug.png")
        except WebDriverException:
            pass
        return False

    except (WebDriverException, StaleElementReferenceException) as exc:
        print(f"RESULT: LinkedIn connection UI failed: {exc}")
        try:
            driver.save_screenshot("logs/connection_exception_debug.png")
        except WebDriverException:
            pass
        return False
    except Exception as exc:
        print(f"RESULT: Connection request failed: {exc}")
        return False

