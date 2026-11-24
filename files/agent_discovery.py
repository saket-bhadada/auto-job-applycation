"""Groq tool-calling agent: finds internships + referral contacts, logs to Excel.

The agent NEVER drafts or sends any message. It only discovers and records contacts.
"""
from __future__ import annotations

import json
import logging
import re
import time
from typing import Any, Callable, TypeVar
from urllib.parse import urlparse

import groq
import httpx
from groq import Groq
from tavily import TavilyClient

import config
from excel_manager import ExcelManager

log = logging.getLogger("discovery")
T = TypeVar("T")

BLOCKED_ROLE = re.compile(
    r"\b(ceo|cto|cfo|coo|cio|chief|founder|co-?founder|president|vp|vice president|"
    r"interns?|internship|trainee|former)\b|\bex-", re.I)
TIER_LABELS = {1: "Tier 1 - Alumni", 2: "Tier 2 - SWE", 3: "Tier 3 - Recruiter"}

GROQ_RETRY = (groq.RateLimitError, groq.APIConnectionError, groq.InternalServerError)
TAVILY_FATAL = {"InvalidAPIKeyError", "BadRequestError", "ForbiddenError"}


def with_backoff(fn: Callable[[], T], *, what: str, retry_on: tuple[type[BaseException], ...],
                 attempts: int = 5, base: float = 2.0, fatal_names: set[str] | None = None) -> T:
    """Exponential backoff; honours a Retry-After header when present."""
    for i in range(attempts):
        try:
            return fn()
        except retry_on as exc:
            if fatal_names and type(exc).__name__ in fatal_names:
                raise
            if i == attempts - 1:
                raise
            wait = base * (2 ** i)
            resp = getattr(exc, "response", None)
            try:
                if resp is not None and resp.headers.get("retry-after"):
                    wait = max(wait, float(resp.headers["retry-after"]) + 0.5)
            except (ValueError, AttributeError):
                pass
            log.warning("%s failed (%s). Retry %d/%d in %.1fs", what, type(exc).__name__, i + 1, attempts - 1, wait)
            time.sleep(wait)
    raise RuntimeError("unreachable")


def _norm_url(url: str) -> str:
    p = urlparse(url.strip())
    host = p.netloc.lower().removeprefix("www.")
    if host.endswith("linkedin.com"):
        host = "linkedin.com"
    return f"{host}{p.path.rstrip('/')}".lower()


TOOLS: list[dict[str, Any]] = [
    {"type": "function", "function": {
        "name": "search_internet",
        "description": "Search the live web (Tavily). Use for job boards and public LinkedIn profile searches.",
        "parameters": {"type": "object", "properties": {
            "query": {"type": "string", "description": "Search query; may use site: operators."}},
            "required": ["query"]}}},
    {"type": "function", "function": {
        "name": "save_to_excel",
        "description": "Save one verified internship plus up to 2 referral contacts. URLs must come from search results.",
        "parameters": {"type": "object", "properties": {
            "company": {"type": "string"},
            "role": {"type": "string", "description": "Exact internship title"},
            "location": {"type": "string", "description": "Exact job location in India (e.g. 'Bengaluru, India', 'Hyderabad', 'Remote - India')"},
            "job_url": {"type": "string"},
            "match_score": {"type": "integer", "description": "0-100 fit with candidate profile"},
            "resume_track": {"type": "string", "enum": ["swe", "aiml", "fullstack", "pm"],
                             "description": "Tailored resume track best matching this role"},
            "primary_name": {"type": "string"}, "primary_role": {"type": "string"},
            "primary_linkedin": {"type": "string"},
            "primary_tier": {"type": "integer", "enum": [1, 2, 3]},
            "backup_name": {"type": "string"}, "backup_role": {"type": "string"},
            "backup_linkedin": {"type": "string"},
            "backup_tier": {"type": "integer", "enum": [1, 2, 3]}},
            "required": ["company", "role", "location", "job_url", "match_score"]}}},
]


class DiscoveryAgent:
    def __init__(self, excel: ExcelManager) -> None:
        self.excel = excel
        self.groq = Groq(api_key=config.GROQ_API_KEY, max_retries=0)
        self.tavily = TavilyClient(api_key=config.TAVILY_API_KEY)
        self.model = config.PRIMARY_MODEL
        self.unavailable_models: set[str] = set()
        self.seen_urls: set[str] = set()
        self.saved = 0
        self.saved_jobs: list[str] = []

    # ---------- tools ----------
    def search_internet(self, query: str) -> dict[str, Any]:
        log.info("search: %s", query)
        res = with_backoff(
            lambda: self.tavily.search(query=query, search_depth="basic", max_results=5),
            what="Tavily search", retry_on=(Exception,), fatal_names=TAVILY_FATAL)
        items = []
        for r in res.get("results", []):
            self.seen_urls.add(_norm_url(r["url"]))
            items.append({"title": (r.get("title") or "")[:100], "url": r["url"],
                          "snippet": (r.get("content") or "")[:280]})
        return {"results": items}

    @staticmethod
    def _verify_india_job(job_url: str, claimed_location: str) -> tuple[bool, str]:
        """Verify job is open and located in India (both claimed location + live ATS page header)."""
        loc = claimed_location.strip()
        if config.INDIA_ONLY and not config.is_india_location(loc):
            return False, (
                f"location '{loc}' rejected: India-only restriction is active "
                "(must be in an Indian city/state or explicitly 'Remote - India')"
            )
        try:
            resp = httpx.get(job_url, follow_redirects=True, timeout=10.0)
            if resp.status_code == 404:
                return False, "job_url returned HTTP 404 (closed listing)"
            if resp.status_code == 200 and config.INDIA_ONLY:
                html = resp.text
                page_loc = ""
                if "lever.co" in job_url.lower():
                    m = re.search(r'class="[^"]*posting-category[^"]*location[^"]*">(.*?)</div>', html, re.S | re.I)
                    if m:
                        page_loc = re.sub(r"<[^>]+>", " ", m.group(1)).strip()
                elif "greenhouse.io" in job_url.lower():
                    m = re.search(r'<div class="location"[^>]*>(.*?)</div>', html, re.S | re.I)
                    if m:
                        page_loc = re.sub(r"<[^>]+>", " ", m.group(1)).strip()
                elif "ashbyhq.com" in job_url.lower():
                    m = re.search(r'"locationName"\s*:\s*"([^"]+)"', html)
                    if m:
                        page_loc = m.group(1).strip()
                if page_loc:
                    if not config.is_india_location(page_loc):
                        return False, f"live job page location is '{page_loc}' (outside India); rejected"
                    loc = page_loc
        except Exception as exc:
            log.debug("Live location check skipped for %s: %s", job_url, exc)
        return True, loc

    def save_to_excel(self, **a: Any) -> dict[str, Any]:
        company, role, job_url = (a.get("company") or "").strip(), (a.get("role") or "").strip(), a.get("job_url") or ""
        claimed_loc = (a.get("location") or "").strip()
        score = int(a.get("match_score") or 0)
        if not (company and role and job_url):
            return {"ok": False, "error": "company, role and job_url are required"}
        if score < config.MIN_MATCH_SCORE:
            return {"ok": False, "error": f"match_score {score} < {config.MIN_MATCH_SCORE}; not saved"}
        if _norm_url(job_url) not in self.seen_urls:
            return {"ok": False, "error": "job_url was not returned by search_internet; use a real search result URL"}

        ok_loc, verified_loc = self._verify_india_job(job_url, claimed_loc)
        if not ok_loc:
            log.info("REJECTED (non-India / closed): %s - %s -> %s", company, role, verified_loc)
            return {"ok": False, "error": verified_loc}

        cells: dict[str, str] = {}
        tiers: list[str] = []
        for prefix, label, col_ref, col_li in (("primary", "Primary", "Primary_Referral", "Primary_LinkedIn"),
                                               ("backup", "Backup", "Backup_Referral", "Backup_LinkedIn")):
            name, rrole = (a.get(f"{prefix}_name") or "").strip(), (a.get(f"{prefix}_role") or "").strip()
            li, tier = (a.get(f"{prefix}_linkedin") or "").strip(), a.get(f"{prefix}_tier")
            if not name and not li:
                continue
            if "linkedin.com/in/" not in li.lower() or _norm_url(li) not in self.seen_urls:
                return {"ok": False, "error": f"{prefix}_linkedin must be a linkedin.com/in/ URL returned by search_internet"}
            if BLOCKED_ROLE.search(rrole) or BLOCKED_ROLE.search(name):
                return {"ok": False, "error": f"{prefix} contact rejected (C-suite/intern/former): {rrole}"}
            if tier not in TIER_LABELS:
                return {"ok": False, "error": f"{prefix}_tier must be 1, 2 or 3"}
            cells[col_ref] = f"{name} - {rrole}" if rrole else name
            cells[col_li] = li
            tiers.append(f"{label}: {TIER_LABELS[tier]}")
        if cells.get("Primary_LinkedIn") and cells.get("Primary_LinkedIn") == cells.get("Backup_LinkedIn"):
            return {"ok": False, "error": "primary and backup must be different people"}

        raw_track = (a.get("resume_track") or "").strip().lower()
        track = raw_track if raw_track in config.RESUME_TRACKS else config.infer_resume_track(role)
        record = {"Company": company, "Role": role, "Location": verified_loc, "Job_URL": job_url,
                  "Match_Score": score, "Resume_Track": track, **cells,
                  "Referral_Tier": " | ".join(tiers) if tiers else "Not found"}
        if not self.excel.add_job(record):
            return {"ok": False, "error": "duplicate Company+Role already in sheet"}
        self.saved += 1
        self.saved_jobs.append(f"{company} - {role} ({verified_loc})")
        log.info("SAVED #%d: %s - %s [%s] (score %d, track=%s, %d referral(s))",
                 self.saved, company, role, verified_loc, score, track, len(tiers))
        return {"ok": True, "saved_so_far": self.saved, "location": verified_loc, "resume_track": track}

    def _dispatch(self, name: str, raw_args: str) -> str:
        try:
            args = json.loads(raw_args or "{}")
            if name == "search_internet":
                out = self.search_internet(args["query"])
            elif name == "save_to_excel":
                out = self.save_to_excel(**args)
            else:
                out = {"error": f"unknown tool {name}"}
        except Exception as exc:  # report to the model so it can adapt
            log.error("Tool %s failed: %s", name, exc)
            out = {"error": f"{type(exc).__name__}: {exc}"}
        return json.dumps(out, ensure_ascii=False)

    # ---------- LLM plumbing ----------
    def _system_prompt(self) -> str:
        p = config.CANDIDATE_PROFILE
        profile = {k: v for k, v in p.items() if k not in {"email", "phone", "resume_path"}}
        return f"""You are an internship discovery and referral-research agent. Today is {time.strftime('%Y-%m-%d')}.
Candidate profile: {json.dumps(profile)}

HARD RULES
- INDIA-ONLY RESTRICTION: ONLY discover and save internships located in INDIA (on-site/hybrid in Indian cities such as Bengaluru/Bangalore, Hyderabad, Pune, Mumbai, Gurugram/Gurgaon, Noida, Delhi NCR, Chennai, Kolkata, Ahmedabad, etc., OR explicitly "Remote - India"). Immediately reject US, Canada, UK, Europe, Singapore, APAC, or worldwide/non-India Remote listings.
- Never draft, suggest or send any email/LinkedIn message. You only discover, verify and log contacts.
- Only use URLs that appeared in search_internet results. Never invent companies, people or URLs.
- Only save jobs that are open, are internships in India, and score >= {config.MIN_MATCH_SCORE}/100 for this candidate.

STEP 1 - JOB SEARCH (INDIA ONLY): run 2-3 broad India-specific searches on ATS job boards like
  `site:boards.greenhouse.io "Intern" ("Bangalore" OR "Bengaluru" OR "Hyderabad" OR "Pune" OR "India") ("Software" OR "Backend" OR "AI")`
  or `site:jobs.lever.co "Intern" ("Bangalore" OR "Bengaluru" OR "India" OR "Gurgaon" OR "Noida" OR "Hyderabad")`
  or `site:jobs.ashbyhq.com "Intern" "India"`.
  Do NOT search for global "Remote" without "India", and do NOT put year numbers like "2027" in quotes.
  Focus: Software Engineering / Backend, AI / ML / Agentic, Full-Stack / Frontend, or Product Management internships matching the request.
  Score each 0-100 on skills, experience, projects, location, and timing. Pick resume_track:
    - "swe": Backend, Distributed Systems, Core Software Engineering
    - "aiml": AI/ML, LLM, RAG, Agentic AI, Deep Learning, Data Science
    - "fullstack": Frontend, React, Full-Stack Web Engineering
    - "pm": Product Management, APM, Technical Product Strategy
STEP 2 - REFERRAL CASCADE (per matched job): run 1-2 targeted `site:linkedin.com/in/` searches for people CURRENTLY at that company:
  Tier 1: Thapar Institute of Engineering & Technology alumni in engineering roles at the company.
  Tier 2: mid-level Software Engineers (SDE-1/SDE-2) matching the job's tech stack.
  Tier 3: University / Campus / Technical Recruiters at the company.
  Fill primary with the best candidate (lowest tier number first), backup with the next best, different person.
  Reject: former/ex employees (check the snippet shows a CURRENT position at that company), C-suite/founders/VPs, and current interns.
  If you cannot verify anyone in 1-2 searches, call save_to_excel immediately with no referral fields.
STEP 3 - LOG IMMEDIATELY: process ONE job at a time (find job -> 1-2 LinkedIn searches -> call save_to_excel right away) before moving to the next job. Stop as soon as the requested number of jobs is saved, and reply with a short plain-text summary."""

    def _trim(self, messages: list[dict[str, Any]], max_chars: int = 9500) -> None:
        """Keep context strictly below Groq's free-tier 7k-8k TPM ceiling while preserving tool-call pairs."""
        # 1. Group messages[2:] into atomic [assistant, tool, tool, ...] turns.
        groups: list[list[dict[str, Any]]] = []
        for m in messages[2:]:
            if m["role"] == "assistant" or not groups:
                groups.append([m])
            else:
                groups[-1].append(m)

        # 2. Compact snippets in all turns except the very latest turn.
        for grp in groups[:-1]:
            for m in grp:
                if m["role"] == "tool" and len(m["content"]) > 260:
                    try:
                        data = json.loads(m["content"])
                        if isinstance(data, dict) and "results" in data:
                            compact = [{"title": r.get("title", "")[:70], "url": r.get("url", "")}
                                       for r in data["results"][:4]]
                            m["content"] = json.dumps({"results": compact}, ensure_ascii=False)
                            continue
                    except Exception:
                        pass
                    m["content"] = m["content"][:240] + "...[trimmed]"

        # 3. Drop oldest complete turn groups if total payload still approaches Groq's 413 limit.
        def _size(grps: list[list[dict[str, Any]]]) -> int:
            return sum(len(json.dumps(x)) for x in messages[:2]) + sum(
                len(json.dumps(m)) for g in grps for m in g
            )

        while len(groups) > 3 and _size(groups) > max_chars:
            groups.pop(0)

        if self.saved_jobs:
            base_sys = self._system_prompt()
            messages[0]["content"] = (
                f"{base_sys}\n\nALREADY SAVED ({self.saved} so far — do NOT re-search or save again): "
                + "; ".join(self.saved_jobs)
            )
        messages[2:] = [m for g in groups for m in g]

    def _chat(self, messages: list[dict[str, Any]]):
        chain = [config.PRIMARY_MODEL, config.FALLBACK_MODEL, *getattr(config, "EXTRA_FALLBACK_MODELS", [])]
        last: Exception | None = None
        for round_idx in range(5):
            models = [m for m in chain if m not in self.unavailable_models]
            if not models:
                break
            for model in models:
                try:
                    resp = self.groq.chat.completions.create(
                        model=model, messages=messages, tools=TOOLS, tool_choice="auto",
                        temperature=0.2, max_completion_tokens=1024,
                    )
                    if model != self.model:
                        log.info("Switched active Groq model -> %s", model)
                        self.model = model
                    return resp.choices[0].message
                except groq.NotFoundError as exc:
                    last = exc
                    self.unavailable_models.add(model)
                    log.warning("Groq[%s] unavailable/decommissioned (404); skipping for rest of run", model)
                except groq.RateLimitError as exc:
                    last = exc
                    log.warning("Groq[%s] hit rate limit (429); failing over immediately", model)
                except groq.APIStatusError as exc:
                    last = exc
                    if exc.status_code == 413:
                        log.warning("Groq[%s] payload too large (413); compacting context", model)
                        self._trim(messages, max_chars=6000)
                        continue
                    log.warning("Groq[%s] status %s: %s", model, exc.status_code, str(exc)[:200])
                except GROQ_RETRY as exc:
                    last = exc
                    log.warning("Groq[%s] transient error (%s); trying backup", model, type(exc).__name__)
                except groq.APIError as exc:
                    last = exc
                    log.warning("Groq[%s] error: %s", model, str(exc)[:200])
            if round_idx < 4:
                wait = 6.0 * (round_idx + 1)
                log.warning("All active Groq models rate-limited; waiting %.1fs (retry %d/4)", wait, round_idx + 1)
                time.sleep(wait)
        raise RuntimeError(f"All Groq models failed: {last}")

    # ---------- main loop ----------
    def run(self, request: str) -> int:
        log.info("Discovery started with model %s", self.model)
        messages: list[dict[str, Any]] = [
            {"role": "system", "content": self._system_prompt()},
            {"role": "user", "content": request}]
        for _ in range(config.MAX_AGENT_STEPS):
            self._trim(messages)
            msg = self._chat(messages)
            calls = msg.tool_calls or []
            if not calls:
                log.info("Agent summary: %s", (msg.content or "").strip()[:800])
                break
            messages.append({"role": "assistant", "content": msg.content or "",
                             "tool_calls": [{"id": c.id, "type": "function",
                                             "function": {"name": c.function.name,
                                                          "arguments": c.function.arguments}} for c in calls]})
            for c in calls:
                messages.append({"role": "tool", "tool_call_id": c.id,
                                 "content": self._dispatch(c.function.name, c.function.arguments)})
        else:
            log.warning("Reached MAX_AGENT_STEPS (%d); stopping.", config.MAX_AGENT_STEPS)
        log.info("Discovery finished: %d new job(s) saved", self.saved)
        return self.saved
