import argparse
import logging
import re
import time

from openpyxl import load_workbook

from config import (
    TRACKER_PATH,
    JOB_SEARCHES,
    EXCLUDED_COMPANIES,
    MIN_RELEVANCE_SCORE,
)
from linkedin.browser import create_driver, search_jobs, LinkedInSessionExpired
from matching.matcher import match_jobs
from linkedin.people import (
    process_people_for_job,
    process_pending_connections_for_job,
    search_people_for_job,
    send_connection_request,
    send_linkedin_message,
    recover_referral_send_lock,
)
from linkedin.conversations import inspect_conversations_for_job
from messages.templates import (
    thank_you_referral_message,
    alternative_process_acknowledgement,
    declined_acknowledgement,
)
from storage.excel import (
    create_tracker,
    upsert_job,
    get_contacts_for_job,
    update_contact_response,
    referral_already_received,
    referral_received,
    get_willing_contacts_except,
    add_message,
    add_activity,
    mark_message_sent,
    upsert_contact,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)


# ============================================================================
# Helpers
# ============================================================================

def normalize(value):
    return "" if value is None else str(value).strip()


def is_excluded_company(company):
    company = normalize(company).lower()
    return any(
        normalize(excluded).lower() in company
        for excluded in EXCLUDED_COMPANIES
        if normalize(excluded)
    )


def job_to_tracker(job):
    """Convert browser.py's lowercase job dictionary to the Excel schema."""
    return {
        "Job ID": job.get("job_id", job.get("Job ID", "")),
        "Company": job.get("company", job.get("Company", "")),
        "Job Title": job.get("title", job.get("Job Title", "")),
        "Job URL": job.get("job_url", job.get("Job URL", "")),
        "Company URL": job.get("company_url", job.get("Company URL", "")),
        "Location": job.get("location", job.get("Location", "")),
        "Posted Date": job.get("posted_date", job.get("Posted Date", "")),
        "Posted Text": job.get("posted_text", job.get("Posted Text", "")),
        "Job Age": job.get("job_age_days", job.get("Job Age", "")),
        "Experience": job.get("experience", job.get("Experience", "")),
        "Min Experience": job.get("min_experience_years", job.get("Min Experience", "")),
        "Max Experience": job.get("max_experience_years", job.get("Max Experience", "")),
        "Employment Type": job.get("employment_type", job.get("Employment Type", "")),
        "Salary": job.get("salary_text", job.get("Salary", "")),
        "Min LPA": job.get("min_lpa", job.get("Min LPA", "")),
        "Max LPA": job.get("max_lpa", job.get("Max LPA", "")),
        "Employees": job.get("employee_text", job.get("Employees", "")),
        "Min Employees": job.get("min_employees", job.get("Min Employees", "")),
        "Max Employees": job.get("max_employees", job.get("Max Employees", "")),
        "AI Score": job.get("ai_score", job.get("AI Score", "")),
        "Full Stack Score": job.get(
            "full_stack_score", job.get("Full Stack Score", "")
        ),
        "Relevance Score": job.get(
            "relevance_score", job.get("Relevance Score", "")
        ),
        "Resume Type": job.get("resume_type", job.get("Resume Type", "")),
        "Resume Name": job.get("resume_name", job.get("Resume Name", "")),
        "Resume Path": job.get("resume_path", job.get("Resume Path", "")),
        "Matched Skills": job.get(
            "matched_skills", job.get("Matched Skills", "")
        ),
        "Search Terms": job.get(
            "search_terms",
            job.get("search_term", job.get("Search Terms", "")),
        ),
        "Application Status": job.get("Application Status", "Not Applied"),
        "Referral Status": job.get("Referral Status", ""),
        "Referral Source": job.get("Referral Source", ""),
        "Referral Source Role": job.get("Referral Source Role", ""),
        "Referral Date": job.get("Referral Date", ""),
    }


# ============================================================================
# Job search
# ============================================================================

def run_job_search(driver):
    """Run every configured search and deduplicate globally by LinkedIn Job ID."""
    all_jobs = {}

    for search_term in JOB_SEARCHES:
        logger.info("Searching LinkedIn jobs for: %s", search_term)

        try:
            jobs = search_jobs(driver, search_term)
        except LinkedInSessionExpired:
            logger.error(
                "LinkedIn session expired/blocked during search '%s'. "
                "Stopping remaining searches instead of recording misleading zero-result searches.",
                search_term,
            )
            raise
        except Exception:
            logger.exception("Search failed: %s", search_term)
            continue

        for job in jobs or []:
            job_id = normalize(
                job.get("job_id", job.get("Job ID", ""))
            )
            if not job_id:
                continue

            if is_excluded_company(
                job.get("company", job.get("Company", ""))
            ):
                continue

            existing_terms = normalize(
                job.get("search_terms", job.get("Search Terms", ""))
            )

            terms = {
                x.strip()
                for x in existing_terms.replace(";", ",").split(",")
                if x.strip()
            }
            terms.add(search_term)
            job["search_terms"] = ", ".join(sorted(terms))

            if job_id not in all_jobs:
                all_jobs[job_id] = job
            else:
                old_terms = normalize(
                    all_jobs[job_id].get("search_terms", "")
                )
                merged = {
                    x.strip()
                    for x in old_terms.replace(";", ",").split(",")
                    if x.strip()
                }
                merged.update(terms)
                all_jobs[job_id]["search_terms"] = ", ".join(sorted(merged))

    logger.info("Unique jobs collected: %d", len(all_jobs))
    return list(all_jobs.values())


def process_jobs(driver, raw_jobs, budget=None):
    """Score jobs, save matches >= threshold, then process people."""
    raw_jobs = [
        job
        for job in (raw_jobs or [])
        if not is_excluded_company(
            job.get("company", job.get("Company", ""))
        )
    ]

    logger.info(
        "Jobs after excluded-company filter: %d",
        len(raw_jobs),
    )

    matched = match_jobs(raw_jobs)
    logger.info(
        "Jobs meeting relevance threshold: %d",
        len(matched),
    )

    if budget is None:
        budget = {
            "connection_requests": 0,
            "referral_messages": 0,
        }

    for number, raw_job in enumerate(matched, start=1):
        company = raw_job.get(
            "company", raw_job.get("Company", "")
        )

        if is_excluded_company(company):
            continue

        job = job_to_tracker(raw_job)
        upsert_job(job)

        job_id = normalize(job.get("Job ID"))
        score = job.get("Relevance Score", 0)

        logger.info(
            "[%d/%d] %s | %s | score=%s",
            number,
            len(matched),
            job.get("Company", ""),
            job.get("Job Title", ""),
            score,
        )

        try:
            actions = process_people_for_job(driver, job, budget)
            logger.info(
                "Outreach actions for %s: %d",
                job_id,
                len(actions),
            )
        except LinkedInSessionExpired:
            logger.error(
                "LinkedIn session expired/blocked while processing %s. Stopping this run.",
                job_id,
            )
            raise
        except Exception:
            logger.exception(
                "People processing failed for %s",
                job_id,
            )


# ============================================================================
# Excel Jobs-sheet reader
# ============================================================================

def load_jobs_from_excel():
    """
    Read the Jobs worksheet directly.

    This is important because the tracker workbook stores Jobs as its
    second worksheet and Tracker() is not the API used by this project.
    """
    create_tracker(TRACKER_PATH)

    workbook = load_workbook(
        TRACKER_PATH,
        read_only=True,
        data_only=True,
    )

    try:
        if "Jobs" not in workbook.sheetnames:
            logger.error(
                "Jobs sheet not found. Sheets: %s",
                workbook.sheetnames,
            )
            return []

        sheet = workbook["Jobs"]

        if sheet.max_row < 2:
            return []

        headers = [
            normalize(cell.value)
            for cell in sheet[1]
        ]

        rows = []

        for values in sheet.iter_rows(
            min_row=2,
            values_only=True,
        ):
            if not any(value not in (None, "") for value in values):
                continue

            row = dict(zip(headers, values))

            if normalize(row.get("Job ID")):
                rows.append(row)

        return rows

    finally:
        workbook.close()


def score_value(job):
    try:
        return float(
            str(job.get("Relevance Score", 0)).replace("%", "").strip()
            or 0
        )
    except (TypeError, ValueError):
        return 0.0


# ============================================================================
# Test outreach
# ============================================================================

def activity_days(contact):
    """Convert visible LinkedIn activity text to approximate days."""
    text = normalize(
        contact.get("Activity", contact.get("activity", ""))
    ).lower()

    if not text:
        return 9999

    if any(
        word in text
        for word in (
            "just now",
            "today",
            "minute",
            "hour",
        )
    ):
        return 0

    match = re.search(r"(\d+)\s*day", text)
    if match:
        return int(match.group(1))

    match = re.search(r"(\d+)\s*week", text)
    if match:
        return int(match.group(1)) * 7

    match = re.search(r"(\d+)\s*month", text)
    if match:
        return int(match.group(1)) * 30

    match = re.search(r"(\d+)\s*year", text)
    if match:
        return int(match.group(1)) * 365

    return 9999


def contact_priority(contact):
    try:
        score = float(
            contact.get(
                "Referral Priority Score",
                contact.get("referral_priority_score", 0),
            )
            or 0
        )
    except (TypeError, ValueError):
        score = 0

    role = normalize(
        contact.get("Role", contact.get("role", ""))
    ).lower()

    # Explicit role ordering.
    if "hiring manager" in role:
        role_score = 100
    elif "engineering manager" in role:
        role_score = 95
    elif "recruiter" in role or "talent acquisition" in role:
        role_score = 90
    elif any(
        x in role
        for x in (
            "principal software engineer",
            "staff software engineer",
        )
    ):
        role_score = 80
    elif "lead software engineer" in role or "tech lead" in role:
        role_score = 78
    elif "senior software engineer" in role:
        role_score = 75
    elif "software engineer" in role:
        role_score = 55
    else:
        role_score = 20

    return role_score, score


def command_test_outreach():
    """
    Test exactly one outreach action using an existing qualifying job
    from the Jobs worksheet.

    It does NOT require a fresh job search first.
    It does NOT bypass CAPTCHA/verification/security controls.
    """
    jobs = [
        job
        for job in load_jobs_from_excel()
        if score_value(job) >= MIN_RELEVANCE_SCORE
    ]

    if not jobs:
        logger.info(
            "No saved job with relevance >= %s found in the Jobs sheet.",
            MIN_RELEVANCE_SCORE,
        )
        return

    # Highest scoring existing job.
    jobs.sort(
        key=lambda job: (
            score_value(job),
            normalize(job.get("Updated", "")),
        ),
        reverse=True,
    )

    job = jobs[0]
    job_id = normalize(job.get("Job ID"))

    print()
    print("ONE-TIME OUTREACH TEST")
    print("======================")
    print(f"Company : {normalize(job.get('Company'))}")
    print(f"Title   : {normalize(job.get('Job Title'))}")
    print(f"Job ID  : {job_id}")
    print(f"Score   : {job.get('Relevance Score')}")
    print(f"Resume  : {normalize(job.get('Resume Type'))}")
    print()

    if referral_already_received(job_id):
        logger.info(
            "Referral already received for Job ID %s. No outreach.",
            job_id,
        )
        return

    driver = create_driver()

    try:
        logger.info("Searching people for this saved job...")

        contacts = search_people_for_job(
            driver,
            job,
            max_people=10,
        )

        if not contacts:
            logger.info(
                "No people were discovered for Job ID %s.",
                job_id,
            )
            return

        existing_contacts = get_contacts_for_job(job_id)

        already_contacted = {
            normalize(
                contact.get("Person", "")
            ).lower()
            for contact in existing_contacts
            if normalize(contact.get("Request Status", "")).lower()
            in {
                "referral requested",
                "connection requested",
                "requested",
                "sent",
                "accepted",
            }
        }

        candidates = []

        for contact in contacts:
            person = normalize(
                contact.get("Person", contact.get("person", ""))
            )

            if not person:
                continue

            if person.lower() in already_contacted:
                continue
            activity_age = activity_days(contact)

            if activity_age != 9999 and activity_age > 30:
                continue

            candidates.append(contact)

        if not candidates:
            logger.info(
                "No uncontacted person with visible activity "
                "within the last 30 days was found."
            )
            return

        candidates.sort(
            key=lambda contact: (
                contact_priority(contact),
                -activity_days(contact),
            ),
            reverse=True,
        )

        contact = candidates[0]

        person = normalize(
            contact.get("Person", contact.get("person", ""))
        )
        role = normalize(
            contact.get("Role", contact.get("role", ""))
        )
        degree = normalize(
            contact.get(
                "Connection Degree",
                contact.get("connection_degree", ""),
            )
        )
        activity = normalize(
            contact.get(
                "Activity",
                contact.get("activity", ""),
            )
        )
        profile_url = normalize(
            contact.get(
                "LinkedIn URL",
                contact.get("linkedin_url", ""),
            )
        )

        print("Selected person")
        print("----------------")
        print(f"Name     : {person}")
        print(f"Role     : {role}")
        print(f"Degree   : {degree}")
        print(f"Activity : {activity}")
        print()

        first_name = person.split()[0] if person else "there"

        referral_message_text = (
            f"Hi {first_name}, I came across the "
            f"{normalize(job.get('Job Title'))} role at "
            f"{normalize(job.get('Company'))}. My experience aligns "
            f"well with this position. If you are comfortable, could "
            f"you please refer me for this role? I can share my resume. "
            f"Thank you!"
        )

        # Exactly ONE outreach action.
        if (
            "1st" in degree.lower()
            or "first" in degree.lower()
        ):
            logger.info(
                "1st-degree contact selected. Sending one referral request."
            )

            sent = send_linkedin_message(
                driver,
                profile_url,
                referral_message_text,
            )

            message_type = "Referral Request"

        else:
            logger.info(
                "Non-1st-degree contact selected. Sending one connection request."
            )

            # Initial outreach is connection-only. The referral message is sent
            # later, only after a live check confirms the connection is accepted.
            sent = send_connection_request(
                driver,
                profile_url,
                note=None,
                message_text=None,
            )

            message_type = "Connection Request"

        add_message(
            job_id=job_id,
            company=normalize(job.get("Company")),
            person=person,
            message_type=message_type,
            message=(
                referral_message_text
                if message_type == "Referral Request"
                else (
                    "Sent with a personalized connection note."
                    if clean(contact.get("Role", contact.get("role", "")))
                    else "Sent without a connection note."
                )
            ),
            resume_type=normalize(job.get("Resume Type")),
            resume_name=normalize(job.get("Resume Name")),
            resume_path=normalize(job.get("Resume Path")),
            job_url=normalize(job.get("Job URL")),
            attachment_required="No",
            status="Sent" if sent else "Failed",
        )

        if sent:
            mark_message_sent(
                job_id,
                person,
                message_type,
            )

            # Persist the selected contact so a connection request can be
            # continued later by process-pending.
            if "Connection Request" in message_type:
                contact["Request Status"] = "Connection Requested"
                contact["Contact Status"] = "Contacted"

            contact["Response Status"] = (
                contact.get("Response Status")
                or contact.get("response_status")
                or "No Response"
            )
            contact["Referral Status"] = (
                contact.get("Referral Status")
                or contact.get("referral_status")
                or "Not Requested"
            )

            try:
                upsert_contact(contact)
            except Exception:
                logger.exception(
                    "Could not persist test-outreach contact %s for Job ID %s",
                    person,
                    job_id,
                )

            add_activity(
                job_id,
                normalize(job.get("Company")),
                person,
                "Connection Request Sent" if "Connection Request" in message_type else "Message Sent",
                (
                    "One-time outreach test; persisted for pending-connection workflow."
                    if "Connection Request" in message_type
                    else f"One-time outreach test: {message_type}"
                ),
            )

            print(f"RESULT: {message_type} sent.")
        else:
            add_activity(
                job_id,
                normalize(job.get("Company")),
                person,
                "Connection Request Failed" if "Connection Request" in message_type else "Message Send Failed",
                f"One-time outreach test: {message_type}",
            )

            print(
                "RESULT: LinkedIn did not allow the normal UI action."
            )

    finally:
        try:
            driver.quit()
        except Exception:
            pass


# ============================================================================
# Pending connection workflow
# ============================================================================

def command_recover_referral_lock(job_id, person):
    """Explicitly recover a referral lock after confirming no message was sent."""
    logger.info(
        "Recovering referral lock for Job ID %s | Person=%s",
        job_id,
        person,
    )

    if recover_referral_send_lock(job_id, person):
        logger.info("Referral lock recovery completed.")
    else:
        logger.error("Referral lock recovery was not completed.")


def command_process_pending():
    """
    Check only connection requests already saved in Contacts.

    This command never sends a new connection invitation.
    """
    jobs = load_jobs_from_excel()

    if not jobs:
        logger.info("No jobs found in the Jobs sheet.")
        return

    logger.info(
        "Checking pending connection workflows for %d saved jobs.",
        len(jobs),
    )

    driver = create_driver()

    # Zero connection budget guarantees this command cannot send a new invite.
    budget = {
        "connection_requests": 0,
        "referral_messages": 0,
    }

    total_actions = 0

    try:
        for job in jobs:
            job_id = normalize(job.get("Job ID"))
            if not job_id:
                continue

            try:
                actions = process_pending_connections_for_job(
                    driver,
                    job,
                    budget,
                )

                total_actions += (
                    actions if isinstance(actions, int) else 0
                )

                logger.info(
                    "Pending workflow actions for %s: %s",
                    job_id,
                    actions,
                )

            except LinkedInSessionExpired:
                logger.error(
                    "LinkedIn session expired/blocked during pending workflow. Stopping this run."
                )
                raise
            except Exception:
                logger.exception(
                    "Pending connection processing failed for Job ID %s",
                    job_id,
                )

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    logger.info(
        "Pending workflow finished. New connection requests sent: %d; "
        "referral messages sent: %d",
        budget["connection_requests"],
        budget["referral_messages"],
    )

    print()
    print(
        "PENDING WORKFLOW COMPLETE | "
        f"new connection requests={budget['connection_requests']} | "
        f"referral messages sent={budget['referral_messages']} | "
        f"actions={total_actions}"
    )


def load_jobs_for_conversations():
    return load_jobs_from_excel()


def command_check_messages():
    jobs = load_jobs_for_conversations()

    if not jobs:
        logger.info("No jobs found in the Jobs sheet.")
        return

    driver = create_driver()

    try:
        for job in jobs:
            job_id = normalize(job.get("Job ID"))
            if not job_id:
                continue

            try:
                contacts = get_contacts_for_job(job_id)

                if not contacts:
                    continue

                results = inspect_conversations_for_job(
                    driver,
                    contacts,
                )

                if not results:
                    continue

                for result in results:
                    person = normalize(result.get("person"))
                    status = normalize(
                        result.get("response_status")
                    )
                    message = normalize(
                        result.get("message")
                    )

                    if not person:
                        continue

                    update_contact_response(
                        job_id,
                        person,
                        status,
                        message,
                    )

                    add_activity(
                        job_id,
                        normalize(job.get("Company")),
                        person,
                        "Conversation Checked",
                        f"Response Status: {status}; Message: {message}",
                    )

            except Exception:
                logger.exception(
                    "Conversation check failed for Job ID %s",
                    job_id,
                )

    finally:
        try:
            driver.quit()
        except Exception:
            pass


def command_status():
    create_tracker(TRACKER_PATH)

    workbook = load_workbook(
        TRACKER_PATH,
        read_only=True,
        data_only=True,
    )

    try:
        print()
        print("LinkedIn Automation Tracker")
        print("============================")

        for sheet_name in workbook.sheetnames:
            count = max(
                workbook[sheet_name].max_row - 1,
                0,
            )
            print(f"{sheet_name}: {count} records")

    finally:
        workbook.close()


# ============================================================================
# Normal run
# ============================================================================

def command_run():
    create_tracker(TRACKER_PATH)

    driver = create_driver()

    # One budget for the complete run.
    budget = {
        "connection_requests": 0,
        "referral_messages": 0,
    }

    try:
        # Continue already-recorded workflows first.
        jobs = load_jobs_from_excel()

        for job in jobs:
            job_id = normalize(job.get("Job ID"))
            if not job_id:
                continue

            try:
                process_pending_connections_for_job(
                    driver,
                    job,
                    budget,
                )
            except LinkedInSessionExpired:
                logger.error(
                    "LinkedIn session expired/blocked during saved-job recovery. Stopping this run."
                )
                raise
            except Exception:
                logger.exception(
                    "Previous workflow processing failed for Job ID %s",
                    job_id,
                )

        raw_jobs = run_job_search(driver)
        process_jobs(driver, raw_jobs, budget)

        logger.info(
            "Automation completed. Connection requests=%d; referral messages=%d.",
            budget["connection_requests"],
            budget["referral_messages"],
        )

    finally:
        try:
            driver.quit()
        except Exception:
            pass


# ============================================================================
# CLI
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="LinkedIn job/referral automation"
    )

    parser.add_argument(
        "command",
        nargs="?",
        default="run",
        choices=(
            "run",
            "test-outreach",
            "process-pending",
            "check-messages",
            "status",
            "recover-referral-lock",
        ),
    )

    parser.add_argument(
        "--job-id",
        dest="job_id",
        help="Job ID used with recover-referral-lock.",
    )
    parser.add_argument(
        "--person",
        dest="person",
        help="Exact person name used with recover-referral-lock.",
    )

    args = parser.parse_args()

    if args.command == "recover-referral-lock":
        if not args.job_id or not args.person:
            parser.error(
                "recover-referral-lock requires --job-id and --person"
            )
        command_recover_referral_lock(args.job_id, args.person)
        return

    if args.command == "run":
        command_run()

    elif args.command == "test-outreach":
        command_test_outreach()

    elif args.command == "process-pending":
        command_process_pending()

    elif args.command == "check-messages":
        command_check_messages()

    elif args.command == "status":
        command_status()


if __name__ == "__main__":
    main()
