import re

from config import (
    PROFILE_AI,
    PROFILE_FULL_STACK,
    MIN_RELEVANCE_SCORE,
    MIN_EXPERIENCE_YEARS,
    MAX_EXPERIENCE_YEARS,
    MAX_JOB_AGE_DAYS,
    AI_JOB_TITLE_KEYWORDS,
    FULL_STACK_JOB_TITLE_KEYWORDS,
    RESUMES,
)


# ============================================================
# TEXT HELPERS
# ============================================================

def normalize(text):
    """
    Normalize text for reliable keyword matching.
    """
    text = "" if text is None else str(text)
    text = text.lower()

    # Keep common technical characters.
    text = re.sub(r"[^a-z0-9+#./ -]", " ", text)

    return re.sub(r"\s+", " ", text).strip()


def _pattern(term):
    """
    Build a reasonably safe whole-term regex pattern.
    """
    term = normalize(term)

    if not term:
        return None

    escaped = re.escape(term).replace(r"\ ", r"\s+")

    return rf"(?<![a-z0-9]){escaped}(?![a-z0-9])"


def contains_term(text, term):
    """
    Check whether a complete term exists in text.
    """
    pattern = _pattern(term)

    return bool(
        pattern and re.search(pattern, normalize(text), re.I)
    )


# ============================================================
# SKILL MATCHING
# ============================================================

def calculate_skill_hits(text, skills):
    """
    Return skills from the resume that appear in the job text.
    """
    normalized_text = normalize(text)

    hits = []

    for skill in skills or []:
        pattern = _pattern(skill)

        if pattern and re.search(pattern, normalized_text, re.I):
            hits.append(skill)

    return sorted(
        set(hits),
        key=lambda x: str(x).lower(),
    )


def calculate_skill_score(hits, total_skills):
    """
    Convert skill overlap into 0-35 points.

    More matched skills increase the score, but we cap the
    contribution so a job cannot become a strong match purely
    because of keyword stuffing.
    """
    hit_count = len(hits)

    if hit_count == 0:
        return 0

    # Progressive scoring.
    if hit_count >= 10:
        return 35

    if hit_count >= 8:
        return 32

    if hit_count >= 6:
        return 28

    if hit_count >= 5:
        return 25

    if hit_count >= 4:
        return 21

    if hit_count >= 3:
        return 17

    if hit_count >= 2:
        return 12

    return 7


# ============================================================
# TITLE / ROLE MATCHING
# ============================================================

def title_matches(title, keywords):
    """
    Backward-compatible helper.

    Returns True when one of the configured keywords is present
    in the normalized job title.
    """
    normalized_title = normalize(title)

    return any(
        normalize(keyword) in normalized_title
        for keyword in keywords or []
    )


def title_role_score(title, profile_type):
    """
    Score the job title from 0-25.

    We intentionally give generic Software Engineer titles less
    weight than titles that explicitly match the candidate's
    target profile.
    """
    title = normalize(title)

    if not title:
        return 0, "No title match"

    if profile_type == PROFILE_AI:

        strong_titles = [
            "ai engineer",
            "ai/ml engineer",
            "ai ml engineer",
            "machine learning engineer",
            "ml engineer",
            "generative ai engineer",
            "genai engineer",
            "applied ai engineer",
            "nlp engineer",
            "ai software engineer",
            "ai backend engineer",
            "machine learning software engineer",
            "python ai engineer",
            "data scientist",
            "data science engineer",
            "machine learning developer",
            "ml developer",
        ]

        medium_titles = [
            "software engineer",
            "python engineer",
            "backend engineer",
            "python developer",
            "data engineer",
        ]

        for keyword in strong_titles:
            if contains_term(title, keyword):
                return 25, f"Strong AI title match: {keyword}"

        for keyword in medium_titles:
            if contains_term(title, keyword):
                return 12, f"Potential AI title match: {keyword}"

        # AI-specific terms inside otherwise unusual titles.
        ai_terms = [
            "artificial intelligence",
            "machine learning",
            "generative ai",
            "genai",
            "nlp",
            "computer vision",
            "llm",
            "deep learning",
        ]

        if any(contains_term(title, term) for term in ai_terms):
            return 20, "AI-related title"

        return 0, "No AI title match"

    # --------------------------------------------------------
    # FULL STACK PROFILE
    # --------------------------------------------------------

    strong_titles = [
        "full stack developer",
        "full-stack developer",
        "full stack engineer",
        "full-stack engineer",
        "full stack software engineer",
        "full-stack software engineer",
        "react engineer",
        "react developer",
        "frontend engineer",
        "front end engineer",
        "backend engineer",
        "back end engineer",
        "python backend engineer",
        "python backend developer",
        "python developer",
    ]

    medium_titles = [
        "software engineer",
        "software developer",
        "web developer",
        "web engineer",
    ]

    for keyword in strong_titles:
        if contains_term(title, keyword):
            return 25, f"Strong Full Stack title match: {keyword}"

    for keyword in medium_titles:
        if contains_term(title, keyword):
            return 12, f"Potential Full Stack title match: {keyword}"

    return 0, "No Full Stack title match"


# ============================================================
# EXPERIENCE MATCHING
# ============================================================

def _num(value):
    """
    Safely convert a value to float.
    """
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _get_experience(job):
    """
    Read experience information from either the internal job
    representation or Excel-style field names.
    """
    minimum = _num(
        job.get(
            "min_experience_years",
            job.get("Min Experience"),
        )
    )

    maximum = _num(
        job.get(
            "max_experience_years",
            job.get("Max Experience"),
        )
    )

    return minimum, maximum


def _experience_points(job):
    """
    Return:

        (points, hard_mismatch, explanation)

    Professional experience target:
        1.5 - 2.0 years

    Jobs requiring materially more than 2 years are treated
    as incompatible.

    Jobs with no stated experience requirement receive a neutral
    score rather than being penalized heavily.
    """
    minimum, maximum = _get_experience(job)

    # Explicitly too senior.
    if minimum is not None and minimum > MAX_EXPERIENCE_YEARS:
        return (
            0,
            True,
            f"Requires {minimum:g}+ years; above target",
        )

    if maximum is not None and maximum < MIN_EXPERIENCE_YEARS:
        return (
            0,
            True,
            f"Maximum {maximum:g} years; below target",
        )

    # No experience information.
    if minimum is None and maximum is None:
        return (
            10,
            False,
            "Experience not specified",
        )

    # Ideal range.
    if (
        minimum is not None
        and minimum <= MIN_EXPERIENCE_YEARS
        and (
            maximum is None
            or maximum >= MAX_EXPERIENCE_YEARS
        )
    ):
        return (
            20,
            False,
            "Experience fits target range",
        )

    # Minimum requirement is within our experience.
    if (
        minimum is not None
        and minimum <= MAX_EXPERIENCE_YEARS
        and (
            maximum is None
            or maximum >= MIN_EXPERIENCE_YEARS
        )
    ):
        return (
            17,
            False,
            "Experience reasonably fits target",
        )

    # Still potentially reasonable.
    return (
        12,
        False,
        "Experience partially fits target",
    )


# ============================================================
# DESCRIPTION RELEVANCE
# ============================================================

def calculate_description_relevance(
    title,
    description,
    hits,
    profile_type,
):
    """
    Award up to 10 additional points when the job description
    contains meaningful profile-specific technology.

    This helps prevent a generic 'Software Engineer' title from
    becoming a strong match based only on the title.
    """
    text = normalize(f"{title} {description}")

    if not text:
        return 0

    score = 0

    if profile_type == PROFILE_AI:

        ai_terms = [
            "machine learning",
            "artificial intelligence",
            "generative ai",
            "genai",
            "llm",
            "langchain",
            "langgraph",
            "rag",
            "natural language processing",
            "nlp",
            "computer vision",
            "tensorflow",
            "scikit-learn",
            "pandas",
            "numpy",
            "model training",
            "model evaluation",
            "deep learning",
        ]

    else:

        ai_terms = [
            "react",
            "typescript",
            "javascript",
            "redux",
            "frontend",
            "front end",
            "full stack",
            "backend",
            "back end",
            "python",
            "fastapi",
            "rest api",
            "restful",
            "aws",
            "docker",
            "sql",
            "mysql",
            "api development",
        ]

    description_hits = 0

    for term in ai_terms:
        if contains_term(text, term):
            description_hits += 1

    if description_hits >= 6:
        score = 10
    elif description_hits >= 4:
        score = 8
    elif description_hits >= 3:
        score = 6
    elif description_hits >= 2:
        score = 4
    elif description_hits >= 1:
        score = 2

    return score


# ============================================================
# FRESHNESS
# ============================================================

def freshness_points(job):
    """
    Give up to 10 points based on posting age.

    Fresh jobs are prioritized but freshness alone cannot make
    a bad job relevant.
    """
    age = _num(
        job.get(
            "job_age_days",
            job.get("Job Age"),
        )
    )

    if age is None:
        return 0

    if age <= 1:
        return 10

    if age <= 2:
        return 9

    if age <= 3:
        return 8

    if age <= 4:
        return 6

    if age <= 5:
        return 4

    if age <= 7:
        return 2

    return 0


# ============================================================
# PROFILE SCORING
# ============================================================

def score_profile(job, profile, title_keywords, profile_type):
    """
    Score one job against one resume profile.

    Maximum score:
        Title fit           25
        Skills              35
        Experience          20
        Description fit     10
        Freshness            10
        --------------------------------
        Total              100
    """

    title = job.get(
        "title",
        job.get("Job Title", ""),
    )

    description = job.get(
        "description",
        job.get("Description", ""),
    )

    terms = job.get(
        "search_terms",
        job.get("Search Terms", ""),
    )

    if isinstance(terms, list):
        terms = " ".join(
            map(str, terms)
        )

    # Include search terms because they tell us why the job was
    # discovered, but they should not dominate the score.
    searchable_text = (
        f"{title} "
        f"{description} "
        f"{terms}"
    )

    hits = calculate_skill_hits(
        searchable_text,
        profile.get("skills", []),
    )

    skill_score = calculate_skill_score(
        hits,
        len(profile.get("skills", [])),
    )

    title_score, title_reason = title_role_score(
        title,
        profile_type,
    )

    experience_score, experience_mismatch, experience_reason = (
        _experience_points(job)
    )

    description_score = calculate_description_relevance(
        title,
        description,
        hits,
        profile_type,
    )

    freshness_score = freshness_points(job)

    total_score = (
        title_score
        + skill_score
        + experience_score
        + description_score
        + freshness_score
    )

    return {
        "score": min(100, total_score),
        "hits": hits,
        "skill_score": skill_score,
        "title_score": title_score,
        "title_reason": title_reason,
        "experience_score": experience_score,
        "experience_reason": experience_reason,
        "description_score": description_score,
        "freshness_score": freshness_score,
        "hard_experience_mismatch": experience_mismatch,
    }


# ============================================================
# JOB SCORING
# ============================================================

def score_job(job):
    """
    Score the job against both resumes and select the stronger
    resume.
    """

    ai = score_profile(
        job,
        RESUMES[PROFILE_AI],
        AI_JOB_TITLE_KEYWORDS,
        PROFILE_AI,
    )

    fs = score_profile(
        job,
        RESUMES[PROFILE_FULL_STACK],
        FULL_STACK_JOB_TITLE_KEYWORDS,
        PROFILE_FULL_STACK,
    )

    # --------------------------------------------------------
    # Select best profile.
    # --------------------------------------------------------

    if ai["score"] >= fs["score"]:
        selected = PROFILE_AI
        selected_result = ai
    else:
        selected = PROFILE_FULL_STACK
        selected_result = fs

    score = selected_result["score"]
    hits = selected_result["hits"]

    resume = RESUMES[selected]

    # --------------------------------------------------------
    # Store scoring information directly on the job.
    # --------------------------------------------------------

    job["ai_score"] = ai["score"]
    job["full_stack_score"] = fs["score"]

    job["relevance_score"] = score

    job["resume_type"] = selected
    job["resume_name"] = resume["name"]
    job["resume_path"] = resume["path"]

    job["matched_skills"] = ", ".join(
        hits[:20]
    )

    job["match_status"] = (
        "MATCH"
        if score >= MIN_RELEVANCE_SCORE
        else "LOW_MATCH"
    )

    job["experience_mismatch"] = (
        "Yes"
        if (
            ai["hard_experience_mismatch"]
            and fs["hard_experience_mismatch"]
        )
        else "No"
    )

    # Useful diagnostic information for Excel/logging later.
    job["ai_title_score"] = ai["title_score"]
    job["full_stack_title_score"] = fs["title_score"]

    job["ai_skill_score"] = ai["skill_score"]
    job["full_stack_skill_score"] = fs["skill_score"]

    job["ai_experience_score"] = ai["experience_score"]
    job["full_stack_experience_score"] = fs["experience_score"]

    job["ai_description_score"] = ai["description_score"]
    job["full_stack_description_score"] = fs["description_score"]

    job["freshness_score"] = selected_result[
        "freshness_score"
    ]

    job["match_reason"] = (
        f"{selected}: "
        f"title={selected_result['title_score']}, "
        f"skills={selected_result['skill_score']}, "
        f"experience={selected_result['experience_score']}, "
        f"description={selected_result['description_score']}, "
        f"freshness={selected_result['freshness_score']}"
    )

    return job


# ============================================================
# MATCH MULTIPLE JOBS
# ============================================================

def match_jobs(jobs):
    """
    Deduplicate jobs by Job ID, score them against both resumes,
    and return only jobs meeting the configured relevance score.
    """

    results = []
    seen = set()

    for job in jobs or []:

        job_id = str(
            job.get(
                "job_id",
                job.get("Job ID", ""),
            )
        ).strip()

        # ----------------------------------------------------
        # Deduplicate.
        # ----------------------------------------------------

        if job_id and job_id in seen:
            continue

        if job_id:
            seen.add(job_id)

        # ----------------------------------------------------
        # Score.
        # ----------------------------------------------------

        scored = score_job(job)

        # ----------------------------------------------------
        # Reject hard experience mismatches.
        # ----------------------------------------------------

        if (
            scored.get("experience_mismatch") == "Yes"
            and scored.get("relevance_score", 0) < MIN_RELEVANCE_SCORE
        ):
            continue

        # ----------------------------------------------------
        # Keep only relevance >= configured threshold.
        # ----------------------------------------------------

        if (
            scored.get("relevance_score", 0)
            >= MIN_RELEVANCE_SCORE
        ):
            results.append(scored)

    # --------------------------------------------------------
    # Sort:
    #
    # 1. Highest relevance
    # 2. Newer jobs first
    # --------------------------------------------------------

    results.sort(
        key=lambda x: (
            x.get("relevance_score", 0),
            -(
                x.get("job_age_days")
                if x.get("job_age_days") is not None
                else 999
            ),
        ),
        reverse=True,
    )

    return results