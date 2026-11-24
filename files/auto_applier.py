"""Excel YES/NO watcher + Playwright form filler."""
from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, sync_playwright

import config
from excel_manager import ExcelLockedError, ExcelManager

log = logging.getLogger("applier")

# Ordered by specificity: first match wins.
FIELD_PATTERNS: list[tuple[str, str]] = [
    ("first_name", r"first.?name|given.?name|fname"),
    ("last_name", r"last.?name|family.?name|surname|lname"),
    ("email", r"e-?mail"),
    ("phone", r"phone|mobile|contact.?number"),
    ("linkedin", r"linkedin"),
    ("github", r"github"),
    ("portfolio", r"portfolio|personal.?(web)?site|website"),
    ("university", r"school|university|college|institution"),
    ("full_name", r"full.?name|your.?name|^name$"),
]
CONFIRM_RE = re.compile(r"thank you for applying|application (has been )?(submitted|received)|"
                        r"successfully submitted|we('ve| have) received your application|thanks for applying", re.I)


class AutoApplier:
    def __init__(self, excel: ExcelManager) -> None:
        self.excel = excel

    # ---------- watcher ----------
    def watch(self) -> None:
        log.info("Watching %s every %ds. Type YES / NO in the Approval column and save. Ctrl+C to stop.",
                 config.EXCEL_PATH.name, config.POLL_INTERVAL_SECONDS)
        try:
            while True:
                try:
                    self.excel.skip_rejected()
                    for row in self.excel.get_approved_waiting():
                        self.apply(row)
                except ExcelLockedError as exc:
                    log.warning("%s - will try again next cycle", exc)
                except Exception:
                    log.exception("Watcher cycle failed")
                time.sleep(config.POLL_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            log.info("Watcher stopped.")

    # ---------- apply one row ----------
    def apply(self, row: dict[str, Any]) -> None:
        jid, label = row["Job_ID"], f"{row['Company']} - {row['Role']}"
        track, resume_path = config.select_resume(
            role=str(row.get("Role", "")),
            track=str(row.get("Resume_Track", "")),
        )
        log.info("Applying: %s [track=%s, resume=%s]", label, track, resume_path.name)
        self.excel.update_row(jid, Apply_Status="In_Progress")
        status = "Manual_Needed"
        try:
            status = self._run_browser(str(row["Job_URL"]), resume_path)
        except Exception:
            log.exception("Browser automation failed for %s", label)
        self.excel.update_row(jid, Apply_Status=status)
        log.info("%s -> %s", label, status)

    def _run_browser(self, url: str, resume_path: Path) -> str:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=False)
            try:
                page = browser.new_context().new_page()
                page.goto(self._form_url(url), wait_until="domcontentloaded", timeout=60_000)
                page.wait_for_timeout(2500)
                self._open_form(page)

                filled = self._fill_fields(page)
                uploaded = self._upload_resume(page, resume_path)
                missing = self._missing_required(page)
                log.info("Filled: %s | resume uploaded: %s | unfilled required: %s",
                         sorted(filled), uploaded, missing or "none")

                log.info("Pausing %ds for visual verification...", config.VERIFY_SECONDS)
                page.wait_for_timeout(config.VERIFY_SECONDS * 1000)

                if uploaded and not missing and config.AUTO_SUBMIT:
                    self._click_submit(page)
                elif missing:
                    log.warning("Custom/required fields left: complete them in the browser (%ds window).",
                                config.MANUAL_HOLD_SECONDS)
                else:
                    log.info("AUTO_SUBMIT off or resume missing: submit manually in the browser.")

                # Polls for a confirmation message; also covers CAPTCHA solved by hand.
                return "Applied" if self._wait_confirmation(page, config.MANUAL_HOLD_SECONDS) else "Manual_Needed"
            finally:
                browser.close()

    # ---------- helpers ----------
    @staticmethod
    def _form_url(url: str) -> str:
        u = url.rstrip("/")
        if "jobs.lever.co" in u and not u.endswith("/apply"):
            return u + "/apply"
        if "jobs.ashbyhq.com" in u and not u.endswith("/application"):
            return u + "/application"
        return url

    @staticmethod
    def _open_form(page: Page) -> None:
        if page.locator("input[type=email], input[name*=email i]").count():
            return
        btn = page.get_by_role("link", name=re.compile(r"^\s*apply( now| for this job)?\s*$", re.I)).first
        if not btn.count():
            btn = page.get_by_role("button", name=re.compile(r"^\s*apply( now| for this job)?\s*$", re.I)).first
        if btn.count():
            btn.click()
            page.wait_for_timeout(2500)

    def _values(self) -> dict[str, str]:
        p_ = config.CANDIDATE_PROFILE
        parts = p_["name"].split()
        return {"first_name": parts[0] if parts else "", "last_name": " ".join(parts[1:]),
                "full_name": p_["name"], "email": p_["email"], "phone": p_["phone"],
                "linkedin": p_["linkedin_url"], "github": p_["github_url"],
                "portfolio": p_["portfolio_url"], "university": p_["college"]}

    def _fill_fields(self, page: Page) -> set[str]:
        values, done = self._values(), set()
        sel = ("input:not([type=hidden]):not([type=file]):not([type=checkbox]):not([type=radio])"
               ":not([type=submit]):not([type=button]), textarea")
        for el in page.locator(sel).all():
            try:
                if not el.is_visible() or el.input_value().strip():
                    continue
                desc = el.evaluate(
                    "e => [e.name, e.id, e.placeholder, e.getAttribute('aria-label'),"
                    " ...(e.labels ? [...e.labels].map(l => l.innerText) : [])].join(' ')")
                for key, pat in FIELD_PATTERNS:
                    if re.search(pat, desc, re.I) and values.get(key) and key not in done:
                        el.fill(values[key])
                        done.add(key)
                        break
            except Exception as exc:
                log.debug("skip field: %s", exc)
        return done

    @staticmethod
    def _upload_resume(page: Page, resume: Path) -> bool:
        if not resume.is_file():
            log.error("Resume not found: %s", resume)
            return False
        inputs = page.locator("input[type=file]").all()
        if not inputs:
            return False
        target = inputs[0]
        for el in inputs:
            try:
                d = el.evaluate("e => [e.name, e.id, e.getAttribute('aria-label')].join(' ')")
                if re.search(r"resume|cv", d, re.I):
                    target = el
                    break
            except Exception:
                pass
        try:
            target.set_input_files(str(resume))
            page.wait_for_timeout(1500)
            return True
        except Exception as exc:
            log.warning("Resume upload failed: %s", exc)
            return False

    @staticmethod
    def _missing_required(page: Page) -> list[str]:
        missing = []
        for el in page.locator("[required], [aria-required='true']").all():
            try:
                if not el.is_visible():
                    continue
                kind = (el.get_attribute("type") or "").lower()
                if kind in {"file", "checkbox", "radio", "hidden"}:
                    continue
                if not el.input_value().strip():
                    missing.append(el.evaluate("e => e.name || e.id || e.getAttribute('aria-label') || e.tagName"))
            except Exception:
                pass
        return missing

    @staticmethod
    def _click_submit(page: Page) -> None:
        btn = page.locator("button[type=submit], input[type=submit]").first
        if not btn.count():
            btn = page.get_by_role("button", name=re.compile(r"submit", re.I)).first
        if btn.count():
            btn.click()
            log.info("Submit clicked.")
        else:
            log.warning("No submit button found.")

    @staticmethod
    def _wait_confirmation(page: Page, seconds: int) -> bool:
        deadline = time.time() + seconds
        while time.time() < deadline:
            try:
                if page.get_by_text(CONFIRM_RE).first.is_visible():
                    return True
            except Exception:
                pass
            page.wait_for_timeout(2000)
        return False
