from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = PROJECT_DIR / "data"
LOG_DIR = PROJECT_DIR / "logs"
RESUMES_DIR = PROJECT_DIR / "resumes"
TRACKER_PATH = DATA_DIR / "linkedin_tracker.xlsx"

CHROME_BINARY = "/opt/google/chrome/chrome"
CHROME_PROFILE = PROJECT_DIR / "chrome_profile"

MIN_COMPANY_EMPLOYEES = 200
MAX_JOB_AGE_DAYS = 7
MIN_EXPERIENCE_YEARS = 1.5
MAX_EXPERIENCE_YEARS = 2.0
MIN_RELEVANCE_SCORE = 60

EXCLUDED_COMPANIES = [
    "Ascentt",
]

SCROLL_WAIT_SECONDS = 2
DETAIL_WAIT_SECONDS = 2
COMPANY_WAIT_SECONDS = 2
PEOPLE_WAIT_SECONDS = 2
MAX_SCROLLS = 5

PAGE_LOAD_TIMEOUT_SECONDS = 35
SCRIPT_TIMEOUT_SECONDS = 20
SEARCH_NAVIGATION_RETRIES = 2

# ---------------------------------------------------------------------------
# Job searches
# ---------------------------------------------------------------------------

JOB_SEARCHES = [
    "Full Stack Developer",
    "Full Stack Engineer",
    "Software Engineer Full Stack",
    "Python Full Stack Developer",
    "Full Stack Python Developer",
    "AI Engineer",
    "AI ML Engineer",
    "Machine Learning Engineer",
    "Generative AI Engineer",
    "GenAI Engineer",
    "Applied AI Engineer",
    "ML Engineer",
    "NLP Engineer",
    "Python Backend Engineer",
    "Backend Engineer Python",
    "AI Software Engineer",
    "AI Backend Engineer",
    "Machine Learning Software Engineer",
    "Python AI Engineer",
    "React Developer",
    "React Engineer",
    "Python Developer",
]

# ---------------------------------------------------------------------------
# People / referral discovery
# ---------------------------------------------------------------------------

PEOPLE_ROLES = [
    "Recruiter",
    "Recruiting",
    "Talent Acquisition",
    "Talent Acquisition Partner",
    "Hiring Manager",
    "Engineering Manager",
    "Principal Engineer",
    "Principal Software Engineer",
    "Staff Engineer",
    "Staff Software Engineer",
    "Lead Engineer",
    "Lead Software Engineer",
    "Senior Software Engineer",
    "Software Engineer",
]

HIGH_VALUE_ROLES = [
    "Recruiter",
    "Talent Acquisition",
    "Talent Acquisition Partner",
    "Hiring Manager",
    "Engineering Manager",
]

VERY_HIGH_VALUE_ROLES = [
    "Hiring Manager",
    "Engineering Manager",
]

# We discover at most this many people from a company/job search.
MAX_PEOPLE_PER_JOB = 15

# Only the best few active people are eligible for new outreach.
MAX_INITIAL_REFERRAL_CONTACTS = 3

# Global safety/budget for one complete automation run.
# Three relevant contacts may be contacted per matched job, up to 12
# connection requests across the complete run.
MAX_CONNECTION_REQUESTS_PER_RUN = 12
MAX_REFERRAL_MESSAGES_PER_RUN = 5

# A person must have a visible recent-activity signal in the people-search
# result. People with no visible activity signal are NOT contacted.
ACTIVE_MAX_DAYS = 30

CHECK_LINKEDIN_CONVERSATIONS = True
CONVERSATION_WAIT_SECONDS = 2
RESUME_ALREADY_SHARED = True

PROFILE_AI = "AI / Data Science"
PROFILE_FULL_STACK = "Full Stack"

AI_JOB_TITLE_KEYWORDS = [
    "AI Engineer",
    "AI/ML Engineer",
    "AI ML Engineer",
    "Machine Learning Engineer",
    "ML Engineer",
    "Generative AI Engineer",
    "GenAI Engineer",
    "Applied AI Engineer",
    "NLP Engineer",
    "AI Software Engineer",
    "AI Backend Engineer",
    "Machine Learning Software Engineer",
    "Python AI Engineer",
]

FULL_STACK_JOB_TITLE_KEYWORDS = [
    "Full Stack Developer",
    "Full Stack Engineer",
    "Software Engineer",
    "Python Full Stack Developer",
    "Full Stack Python Developer",
    "React Developer",
    "React Engineer",
    "Python Developer",
    "Backend Engineer",
    "Python Backend Engineer",
    "Backend Engineer Python",
]

RESUMES = {
    PROFILE_AI: {
        "name": "Jaya_Singh_Chauhan_DataScience_AI.pdf",
        "path": str(RESUMES_DIR / "Jaya_Singh_Chauhan_DataScience_AI.pdf"),
        "skills": [
            "Python", "SQL", "Pandas", "NumPy", "Scikit-learn",
            "TensorFlow", "OpenCV", "NLP", "Random Forest",
            "Logistic Regression", "EDA", "Data Cleaning",
            "Feature Engineering", "LangChain", "LangGraph",
            "RAG", "Prompt Engineering", "Generative AI",
            "LLM", "FastAPI", "Flask", "REST API", "OAuth2", "JWT",
            "Pydantic", "MySQL", "Vector DB", "React", "Redux",
            "Docker", "AWS", "EFS", "EKS", "Linux", "Git", "GitHub",
        ],
    },
    PROFILE_FULL_STACK: {
        "name": "Jaya_Singh_Chauhan_Full_Stack.pdf",
        "path": str(RESUMES_DIR / "Jaya_Singh_Chauhan_Full_Stack.pdf"),
        "skills": [
            "React", "TypeScript", "JavaScript", "Redux Toolkit",
            "Redux", "Material UI", "Tailwind", "HTML", "CSS",
            "TanStack Query", "Python", "FastAPI", "Flask", "Node",
            "REST API", "Pydantic", "OAuth2", "JWT", "MySQL", "SQL",
            "AWS", "EFS", "EKS", "Docker", "Linux", "SSH", "SFTP",
            "Grafana", "InfluxDB", "Git", "GitHub", "SSO", "CSP",
            "LangChain", "LangGraph", "RAG", "Vector DB",
        ],
    },
}

RESPONSE_STATUS_NO_RESPONSE = "No Response"
RESPONSE_STATUS_WILLING = "Willing to Refer"
RESPONSE_STATUS_RECEIVED = "Referral Received"
RESPONSE_STATUS_DECLINED = "Declined"
RESPONSE_STATUS_ALTERNATIVE = "Alternative Process"
RESPONSE_STATUS_AMBIGUOUS = "Needs Review"

CONTACT_STATUS_DISCOVERED = "Discovered"
CONTACT_STATUS_CONTACTED = "Contacted"
REQUEST_NOT_REQUESTED = "Not Requested"
REQUEST_CONNECTION = "Connection Requested"
REQUEST_ACCEPTED = "Accepted"
REQUEST_REFERRAL = "Referral Requested"

REFERRAL_NOT_REQUESTED = "Not Requested"
REFERRAL_REQUESTED = "Requested"
REFERRAL_OFFERED = "Offered"
REFERRAL_RECEIVED = "Received"
REFERRAL_DECLINED = "Declined"
REFERRAL_ALTERNATIVE = "Alternative Process"
