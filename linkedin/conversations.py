import re
import time
from typing import Dict, List, Optional

from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import PEOPLE_WAIT_SECONDS


# ---------------------------------------------------------------------------
# Response classification
# ---------------------------------------------------------------------------

WILLING_PATTERNS = [
    r"\bcan refer\b",
    r"\bwill refer\b",
    r"\bi can refer\b",
    r"\bi'll refer\b",
    r"\bi will refer\b",
    r"\bsure[, ]+i can\b",
    r"\bsure[, ]+send\b",
    r"\bhappy to refer\b",
    r"\bhappy to help\b",
    r"\bglad to refer\b",
    r"\bglad to help\b",
    r"\bi can help\b",
    r"\bi'd be happy to refer\b",
    r"\bwould be happy to refer\b",
    r"\bsend me your resume\b",
    r"\bsend your resume\b",
    r"\bsend me the resume\b",
    r"\bshare your resume\b",
    r"\bshare the resume\b",
    r"\bsend me your cv\b",
    r"\bsend your cv\b",
]

REFERRAL_RECEIVED_PATTERNS = [
    r"\bi referred you\b",
    r"\bi have referred you\b",
    r"\bi've referred you\b",
    r"\breferred you\b",
    r"\breferral is done\b",
    r"\bi submitted your referral\b",
    r"\bsubmitted your referral\b",
    r"\breferral submitted\b",
    r"\breferral has been submitted\b",
    r"\bdone[, ]+referred\b",
    r"\bdone[, ]+i referred\b",
    r"\bjust referred you\b",
    r"\breferral has been done\b",
    r"\bcompleted the referral\b",
]

DECLINED_PATTERNS = [
    r"\bi can't refer\b",
    r"\bi cannot refer\b",
    r"\bcan't refer\b",
    r"\bcannot refer\b",
    r"\bnot able to refer\b",
    r"\bunable to refer\b",
    r"\bcan't help with a referral\b",
    r"\bcannot help with a referral\b",
    r"\bnot my team\b",
    r"\bnot in my team\b",
    r"\bnot the right person\b",
    r"\bwrong person\b",
    r"\bi don't handle hiring\b",
    r"\bi do not handle hiring\b",
    r"\bnot involved in hiring\b",
    r"\bno openings\b",
    r"\bno position\b",
]

ALTERNATIVE_PATTERNS = [
    r"\bapply through\b",
    r"\bapply here\b",
    r"\bapply on the website\b",
    r"\bcareers page\b",
    r"\bcareer portal\b",
    r"\bjob portal\b",
    r"\buse this link\b",
    r"\buse the link\b",
    r"\bplease apply\b",
    r"\byou should apply\b",
    r"\bapply directly\b",
    r"\bsubmit through\b",
    r"\bcompany portal\b",
    r"\bexternal application\b",
]

LOW_INFORMATION_PATTERNS = [
    r"^thanks$",
    r"^thank you$",
    r"^okay$",
    r"^ok$",
    r"^sure$",
    r"^noted$",
    r"^cool$",
    r"^great$",
    r"^welcome$",
    r"^hi$",
    r"^hello$",
]


def normalize_text(text: str) -> str:
    if not text:
        return ""

    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def classify_response(text: str) -> str:
    """
    Classify the latest meaningful response.

    Priority:
        Referral Received
        Willing to Refer
        Declined
        Alternative Process
        Needs Review
    """

    text = normalize_text(text).lower()

    if not text:
        return "No Response"

    # Referral completion has the highest priority.
    for pattern in REFERRAL_RECEIVED_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Referral Received"

    # Explicit willingness to refer.
    for pattern in WILLING_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Willing to Refer"

    # Explicit decline.
    for pattern in DECLINED_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Declined"

    # Alternative application route.
    for pattern in ALTERNATIVE_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Alternative Process"

    # Short acknowledgements are not evidence of willingness.
    for pattern in LOW_INFORMATION_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Needs Review"

    return "Needs Review"


# ---------------------------------------------------------------------------
# LinkedIn messaging UI helpers
# ---------------------------------------------------------------------------

def open_messaging(driver):
    """
    Open LinkedIn messaging.

    We intentionally use the normal LinkedIn UI. This does not bypass
    verification, CAPTCHA, login checks, or other LinkedIn protections.
    """

    try:
        driver.get("https://www.linkedin.com/messaging/")
        time.sleep(PEOPLE_WAIT_SECONDS)

        return True

    except WebDriverException:
        return False


def _safe_text(element) -> str:
    try:
        return normalize_text(element.text)
    except (StaleElementReferenceException, WebDriverException):
        return ""


def _safe_attribute(element, attribute: str) -> str:
    try:
        return normalize_text(element.get_attribute(attribute) or "")
    except (StaleElementReferenceException, WebDriverException):
        return ""


def _get_visible_text(driver) -> str:
    try:
        return normalize_text(driver.find_element(By.TAG_NAME, "body").text)
    except WebDriverException:
        return ""


def _find_search_box(driver):
    selectors = [
        "input[placeholder*='Search messages']",
        "input[placeholder*='Search']",
        "input[aria-label*='Search messages']",
        "input[aria-label*='Search']",
        "input[role='combobox']",
    ]

    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)

            for element in elements:
                try:
                    if element.is_displayed() and element.is_enabled():
                        return element
                except WebDriverException:
                    continue

        except WebDriverException:
            continue

    return None


def search_conversation(driver, person_name: str) -> bool:
    """
    Search for a person's conversation using the LinkedIn messaging UI.
    """

    search_box = _find_search_box(driver)

    if search_box is None:
        return False

    try:
        search_box.clear()
        search_box.send_keys(person_name)
        time.sleep(2)

        # Click the first likely result.
        result_selectors = [
            "a[href*='/messaging/thread/']",
            "div[role='button']",
            "li",
            "a",
        ]

        for selector in result_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)

                for element in elements:
                    try:
                        if not element.is_displayed():
                            continue

                        text = _safe_text(element)

                        if (
                            person_name.lower() in text.lower()
                            and len(text) < 500
                        ):
                            driver.execute_script(
                                "arguments[0].click();",
                                element,
                            )

                            time.sleep(1.5)
                            return True

                    except (
                        StaleElementReferenceException,
                        WebDriverException,
                    ):
                        continue

            except WebDriverException:
                continue

        return False

    except WebDriverException:
        return False


# ---------------------------------------------------------------------------
# Conversation extraction
# ---------------------------------------------------------------------------

def extract_conversation_messages(driver) -> List[Dict]:
    """
    Extract message-like elements from the currently opened LinkedIn
    conversation.

    LinkedIn changes its DOM periodically, therefore several selectors
    are intentionally used.
    """

    selectors = [
        "div.msg-s-message-list-content",
        "div.msg-s-event-listitem",
        "li.msg-s-message-list__event",
        "div[class*='msg-s-event-listitem']",
        "li[class*='msg-s-message-list']",
        "div[role='listitem']",
    ]

    found = []

    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)

            for element in elements:
                try:
                    text = _safe_text(element)

                    if not text:
                        continue

                    if len(text) > 3000:
                        continue

                    found.append(
                        {
                            "text": text,
                            "element": element,
                        }
                    )

                except (
                    StaleElementReferenceException,
                    WebDriverException,
                ):
                    continue

        except WebDriverException:
            continue

    # Deduplicate by text.
    unique = []
    seen = set()

    for item in found:
        key = item["text"]

        if key in seen:
            continue

        seen.add(key)
        unique.append(item)

    return unique


def _identify_sender(element) -> Optional[str]:
    selectors = [
        "[data-anonymize='person-name']",
        "a[href*='/in/']",
        "span[dir='ltr']",
        "span",
    ]

    for selector in selectors:
        try:
            elements = element.find_elements(By.CSS_SELECTOR, selector)

            for child in elements:
                text = _safe_text(child)

                if text and len(text) < 150:
                    return text

        except WebDriverException:
            continue

    return None


def extract_latest_received_message(
    driver,
    person_name: str,
) -> Optional[Dict]:
    """
    Return the latest meaningful message that appears to be from the
    contacted person.

    This function deliberately does not assume that every message
    containing the person's name is their message. It uses nearby sender
    metadata when available and falls back to conversation content.
    """

    messages = extract_conversation_messages(driver)

    if not messages:
        return None

    candidate_messages = []

    for item in messages:
        text = item["text"]
        element = item["element"]

        sender = _identify_sender(element)

        # Ignore messages that are clearly from the logged-in account.
        try:
            classes = (element.get_attribute("class") or "").lower()
            aria = (element.get_attribute("aria-label") or "").lower()
            data_test = (element.get_attribute("data-test-id") or "").lower()
            outgoing_markers = ("outgoing", "from-me", "msg-s-event-listitem--outgoing")
            if any(marker in classes for marker in outgoing_markers):
                continue
            if "you" in aria and "message" in aria:
                continue
            if "outgoing" in data_test:
                continue
        except WebDriverException:
            pass

        if sender:
            sender_lower = sender.lower()
            person_lower = person_name.lower()

            if person_lower not in sender_lower:
                continue

        # Ignore obvious UI-only elements.
        lowered = text.lower()

        if lowered in {
            "send",
            "message",
            "type a message",
            "write a message",
        }:
            continue

        candidate_messages.append(
            {
                "sender": sender,
                "text": text,
            }
        )

    if candidate_messages:
        return candidate_messages[-1]

    # Fallback: use the last message-like block.
    last = messages[-1]

    return {
        "sender": None,
        "text": last["text"],
    }


# ---------------------------------------------------------------------------
# Per-contact inspection
# ---------------------------------------------------------------------------

def inspect_person_conversation(
    driver,
    person_name: str,
) -> Dict:
    """
    Inspect one person's LinkedIn conversation.

    Returns:
        {
            "person": ...,
            "response_status": ...,
            "message": ...,
            "found_conversation": bool,
        }
    """

    result = {
        "person": person_name,
        "response_status": "No Response",
        "message": "",
        "found_conversation": False,
    }

    if not search_conversation(driver, person_name):
        return result

    result["found_conversation"] = True

    latest = extract_latest_received_message(
        driver,
        person_name,
    )

    if not latest:
        return result

    text = normalize_text(latest.get("text", ""))

    if not text:
        return result

    result["message"] = text
    result["response_status"] = classify_response(text)

    return result


# ---------------------------------------------------------------------------
# Job-level conversation processing
# ---------------------------------------------------------------------------

def inspect_conversations_for_job(
    driver,
    contacts: List[Dict],
) -> List[Dict]:
    """
    Inspect LinkedIn conversations for all contacts associated with a job.

    Only contacts that have previously been contacted are inspected.
    """

    results = []

    if not contacts:
        return results

    if not open_messaging(driver):
        return results

    for contact in contacts:
        person_name = contact.get("Person", "").strip()

        if not person_name:
            continue

        request_status = str(
            contact.get("Request Status", "")
        ).strip().lower()

        contact_status = str(
            contact.get("Contact Status", "")
        ).strip().lower()

        # Inspect only people whom the automation actually contacted
        # (or whose connection request was accepted). Merely discovering
        # a person must never create a "reply" state.
        contacted_states = {
            "referral requested",
            "connection requested",
            "connection sent",
            "requested",
            "sent",
            "accepted",
            "contacted",
        }
        if request_status not in contacted_states and contact_status != "contacted":
            continue

        result = inspect_person_conversation(
            driver,
            person_name,
        )

        result["job_id"] = contact.get("Job ID", "")
        result["company"] = contact.get("Company", "")
        result["role"] = contact.get("Role", "")
        result["linkedin_url"] = contact.get("LinkedIn URL", "")

        results.append(result)

    return results