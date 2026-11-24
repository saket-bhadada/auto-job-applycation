"""Central configuration: env keys, models, paths, multi-resume routing, and candidate profile."""
from __future__ import annotations

import os
import re
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", override=True)

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip().strip('"').strip("'")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY", "").strip().strip('"').strip("'")

# Primary Groq model with automatic fallback chain when rate-limited or retired.
PRIMARY_MODEL = "openai/gpt-oss-120b"
FALLBACK_MODEL = "llama-3.3-70b-versatile"
EXTRA_FALLBACK_MODELS = ["openai/gpt-oss-20b", "qwen/qwen3.8-27b"]

EXCEL_PATH = BASE_DIR / "internship_tracker.xlsx"
POLL_INTERVAL_SECONDS = 10
MIN_MATCH_SCORE = 75
MAX_AGENT_STEPS = 60
INDIA_ONLY = os.getenv("INDIA_ONLY", "true").lower() in {"1", "true", "yes"}

INDIA_LOCATION_RE = re.compile(
    r"\b("
    r"india|bharat|bengaluru|bangalore|hyderabad|secunderabad|pune|mumbai|navi\s*mumbai|thane|"
    r"gurugram|gurgaon|noida|greater\s*noida|delhi|new\s*delhi|ncr|chennai|kolkata|ahmedabad|"
    r"gandhinagar|jaipur|chandigarh|mohali|panchkula|indore|kochi|cochin|trivandrum|"
    r"thiruvananthapuram|coimbatore|mysuru|mysore|bhubaneswar|visakhapatnam|vizag|surat|"
    r"vadodara|nagpur|lucknow|patiala|dehradun|bhopal|guwahati|"
    r"karnataka|maharashtra|telangana|tamil\s*nadu|haryana|uttar\s*pradesh|gujarat|kerala|"
    r"west\s*bengal|rajasthan|punjab|andhra\s*pradesh|odisha"
    r")\b",
    re.I,
)


def is_india_location(location: str) -> bool:
    """Return True only if location explicitly specifies India or an Indian city/state."""
    loc = (location or "").strip()
    if not loc:
        return False
    return bool(INDIA_LOCATION_RE.search(loc))

# Excel save retries when the file is open in Microsoft Excel.
SAVE_RETRIES = 6
SAVE_RETRY_DELAY = 5.0

# Playwright behaviour
AUTO_SUBMIT = os.getenv("AUTO_SUBMIT", "true").lower() in {"1", "true", "yes"}
VERIFY_SECONDS = int(os.getenv("VERIFY_SECONDS", "15"))
MANUAL_HOLD_SECONDS = int(os.getenv("MANUAL_HOLD_SECONDS", "120"))

PLACEHOLDER = "FILL_ME"


def _resolve(path_str: str) -> Path:
    cleaned = path_str.strip().strip('"').strip("'")
    p = Path(cleaned)
    return p if p.is_absolute() else BASE_DIR / p


DEFAULT_RESUME = _resolve(os.getenv("RESUME_PATH", "resumes/Saket_Bhadada_SWE.pdf"))

RESUME_TRACKS: dict[str, Path] = {
    "swe": _resolve(os.getenv("RESUME_SWE_PATH", str(DEFAULT_RESUME))),
    "aiml": _resolve(os.getenv("RESUME_AIML_PATH", "resumes/Saket_Bhadada_AIML.pdf")),
    "fullstack": _resolve(os.getenv("RESUME_FULLSTACK_PATH", "resumes/Saket_Bhadada_FullStack.pdf")),
    "pm": _resolve(os.getenv("RESUME_PM_PATH", "resumes/Saket_Bhadada_PM.docx")),
}

_TRACK_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("pm", re.compile(r"\b(product\s*manag|apm|associate\s*product|product\s*intern|product\s*strategy|business\s*analyst)\b", re.I)),
    ("aiml", re.compile(r"\b(ai|ml|machine\s*learning|deep\s*learning|llm|genai|generative\s*ai|nlp|agentic|rag|data\s*scien)\b", re.I)),
    ("fullstack", re.compile(r"\b(front-?end|full-?stack|react|ui\s*engineer|web\s*developer|web\s*engineer)\b", re.I)),
]


def infer_resume_track(role: str) -> str:
    """Infer the best resume track ('swe', 'aiml', 'fullstack', 'pm') from a job title."""
    for track, pattern in _TRACK_PATTERNS:
        if pattern.search(role or ""):
            return track
    return "swe"


def select_resume(role: str = "", track: str = "") -> tuple[str, Path]:
    """Return (track_name, existing_resume_path) with graceful fallback."""
    chosen = (track or "").strip().lower()
    if chosen not in RESUME_TRACKS:
        chosen = infer_resume_track(role)
    candidate_path = RESUME_TRACKS[chosen]
    if candidate_path.is_file():
        return chosen, candidate_path
    if DEFAULT_RESUME.is_file():
        return "swe", DEFAULT_RESUME
    for fallback_track, path in RESUME_TRACKS.items():
        if path.is_file():
            return fallback_track, path
    return chosen, candidate_path


CANDIDATE_PROFILE: dict[str, str] = {
    "name": os.getenv("CANDIDATE_NAME", PLACEHOLDER),
    "email": os.getenv("CANDIDATE_EMAIL", PLACEHOLDER),
    "phone": os.getenv("CANDIDATE_PHONE", PLACEHOLDER),
    "college": "Thapar Institute of Engineering & Technology",
    "degree": "B.Tech in Computer Science & Business Systems",
    "graduation_year": os.getenv("CANDIDATE_GRAD_YEAR", PLACEHOLDER),
    "cgpa": os.getenv("CANDIDATE_CGPA", PLACEHOLDER),
    "skills": os.getenv(
        "CANDIDATE_SKILLS",
        "Python, C/C++, JavaScript (ES6+), SQL, FastAPI, Node.js, Express.js, React, "
        "PostgreSQL (pgvector), WebSockets, PyTorch (LSTM/RNN), LangGraph, RAG, LLM Integration",
    ),
    "experience": (
        "Frontend Intern at Motion Education Pvt. Ltd. (React UI & async REST APIs for 5,000+ active users); "
        "Media Head at Thapar Mathematical Club"
    ),
    "projects": (
        "Super Docs (Agentic Document Auditing Platform: FastAPI, LangGraph, React, pgvector, RAG); "
        "Arc-Stream (AI DJ & Energy-Curve Playlist Engine: PyTorch LSTM, FastAPI, Node.js, pgvector HNSW over 114k+ tracks); "
        "TEDxTIET Platform (React, Node.js, Express, PostgreSQL, 99.9% uptime); "
        "Cascade (Supply Chain Capacity & Anomaly Monitoring: RNN, FastAPI, React, PostgreSQL); "
        "Project Pulse (Real-Time Collaboration Platform: React, Node.js, PostgreSQL, WebSockets)"
    ),
    "linkedin_url": os.getenv("CANDIDATE_LINKEDIN", ""),
    "github_url": os.getenv("CANDIDATE_GITHUB", ""),
    "portfolio_url": os.getenv("CANDIDATE_PORTFOLIO", ""),
    "resume_path": str(DEFAULT_RESUME),
}


def validate(for_apply: bool) -> list[str]:
    """Return a list of configuration problems (empty list = OK)."""
    problems: list[str] = []
    if not GROQ_API_KEY:
        problems.append("GROQ_API_KEY missing in .env")
    if not TAVILY_API_KEY:
        problems.append("TAVILY_API_KEY missing in .env")
    if for_apply:
        for key in ("name", "email", "phone"):
            if CANDIDATE_PROFILE[key] in ("", PLACEHOLDER):
                problems.append(f"CANDIDATE_{key.upper()} missing in .env")
        _, resolved_resume = select_resume()
        if not resolved_resume.is_file():
            problems.append(f"Resume not found at {resolved_resume}")
    return problems
