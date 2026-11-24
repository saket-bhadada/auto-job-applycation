"""internship_tracker.xlsx: database + human-in-the-loop control panel."""
from __future__ import annotations

import hashlib
import logging
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import config

log = logging.getLogger("excel")

COLUMNS = [
    "Job_ID", "Company", "Role", "Location", "Job_URL", "Match_Score", "Resume_Track",
    "Primary_Referral", "Primary_LinkedIn", "Referral_Tier",
    "Backup_Referral", "Backup_LinkedIn",
    "Approval", "Apply_Status", "Last_Updated",
]
WIDTHS = {
    "Job_ID": 14, "Company": 22, "Role": 34, "Location": 22, "Job_URL": 48,
    "Match_Score": 13, "Resume_Track": 14,
    "Primary_Referral": 34, "Primary_LinkedIn": 44, "Referral_Tier": 38,
    "Backup_Referral": 34, "Backup_LinkedIn": 44,
    "Approval": 12, "Apply_Status": 16, "Last_Updated": 20,
}
MAX_ROWS = 2000


class ExcelLockedError(RuntimeError):
    """The workbook stayed locked (open in Excel) through all retries."""


def make_job_id(company: str, role: str) -> str:
    """MD5 of Company + Role (normalised); used as the dedupe key."""
    raw = f"{company.strip().lower()}{role.strip().lower()}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:12]


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class ExcelManager:
    def __init__(self, path: Path = config.EXCEL_PATH) -> None:
        self.path = Path(path)
        self._lock = threading.RLock()
        if not self.path.exists():
            self._create()

    # ---------- low-level I/O with PermissionError handling ----------
    def _retry(self, fn, what: str):
        for attempt in range(1, config.SAVE_RETRIES + 1):
            try:
                return fn()
            except PermissionError:
                log.warning("%s: '%s' is open in Excel - close it (retry %d/%d in %.0fs)",
                            what, self.path.name, attempt, config.SAVE_RETRIES, config.SAVE_RETRY_DELAY)
                time.sleep(config.SAVE_RETRY_DELAY)
        raise ExcelLockedError(f"Could not {what} {self.path.name}: file is locked by Excel")

    def _load(self):
        return self._retry(lambda: load_workbook(self.path), "read")

    def _save(self, wb) -> None:
        self._retry(lambda: wb.save(self.path), "save")

    # ---------- workbook creation ----------
    def _create(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Internships"
        ws.append(COLUMNS)
        head_fill = PatternFill("solid", fgColor="1F3864")
        thin = Side(style="thin", color="BFBFBF")
        for idx, name in enumerate(COLUMNS, start=1):
            c = ws.cell(row=1, column=idx)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            ws.column_dimensions[get_column_letter(idx)].width = WIDTHS[name]
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

        track_col = get_column_letter(COLUMNS.index("Resume_Track") + 1)
        dv_track = DataValidation(type="list", formula1='"swe,aiml,fullstack,pm"', allow_blank=False,
                                  showErrorMessage=True, errorTitle="Invalid Resume Track",
                                  error="Choose swe, aiml, fullstack or pm")
        ws.add_data_validation(dv_track)
        dv_track.add(f"{track_col}2:{track_col}{MAX_ROWS}")

        col = get_column_letter(COLUMNS.index("Approval") + 1)
        dv = DataValidation(type="list", formula1='"Pending,YES,NO"', allow_blank=False,
                            showErrorMessage=True, errorTitle="Invalid",
                            error="Choose Pending, YES or NO")
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{MAX_ROWS}")
        rng = f"{col}2:{col}{MAX_ROWS}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"YES"'],
                                      fill=PatternFill("solid", bgColor="C6EFCE")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"NO"'],
                                      fill=PatternFill("solid", bgColor="FFC7CE")))
        self._save(wb)
        log.info("Created %s", self.path.name)

    @staticmethod
    def _headers(ws) -> list[str]:
        hdrs = [str(ws.cell(row=1, column=i).value or "").strip() for i in range(1, ws.max_column + 1)]
        return [h for h in hdrs if h] or list(COLUMNS)

    # ---------- reads ----------
    def get_rows(self) -> list[dict[str, Any]]:
        with self._lock:
            wb = self._load()
            ws = wb.active
            hdrs = self._headers(ws)
            rows = []
            for r in range(2, ws.max_row + 1):
                vals = {c: ws.cell(row=r, column=i + 1).value for i, c in enumerate(hdrs)}
                if vals.get("Job_ID"):
                    vals["_row"] = r
                    rows.append(vals)
            wb.close()
            return rows

    def job_exists(self, company: str, role: str) -> bool:
        jid = make_job_id(company, role)
        return any(r["Job_ID"] == jid for r in self.get_rows())

    @staticmethod
    def _norm(v: Any) -> str:
        return str(v or "").strip().upper()

    def get_approved_waiting(self) -> list[dict[str, Any]]:
        return [r for r in self.get_rows()
                if self._norm(r["Approval"]) == "YES" and self._norm(r["Apply_Status"]) == "WAITING"]

    # ---------- writes ----------
    def add_job(self, record: dict[str, Any]) -> bool:
        """Append a job. Returns False if Company+Role already exists."""
        with self._lock:
            jid = make_job_id(record["Company"], record["Role"])
            wb = self._load()
            ws = wb.active
            hdrs = self._headers(ws)
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value == jid:
                    wb.close()
                    return False
            role_val = record["Role"]
            if "Location" not in hdrs and record.get("Location"):
                role_val = f"{role_val} ({record['Location']})"
            row = {**record, "Role": role_val, "Job_ID": jid,
                   "Resume_Track": record.get("Resume_Track") or config.infer_resume_track(record["Role"]),
                   "Approval": "Pending",
                   "Apply_Status": "Waiting", "Last_Updated": _now()}
            ws.append([row.get(c, "") for c in hdrs])
            for idx in range(1, len(hdrs) + 1):
                ws.cell(row=ws.max_row, column=idx).alignment = Alignment(vertical="top", wrap_text=True)
            self._save(wb)
            return True

    def update_row(self, job_id: str, **fields: Any) -> bool:
        """Re-reads the file right before writing so manual edits are not clobbered."""
        with self._lock:
            wb = self._load()
            ws = wb.active
            hdrs = self._headers(ws)
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value == job_id:
                    for k, v in fields.items():
                        if k in hdrs:
                            ws.cell(row=r, column=hdrs.index(k) + 1, value=v)
                    if "Last_Updated" in hdrs:
                        ws.cell(row=r, column=hdrs.index("Last_Updated") + 1, value=_now())
                    self._save(wb)
                    return True
            wb.close()
            return False

    def skip_rejected(self) -> int:
        """Approval == NO and still Waiting -> Skipped."""
        n = 0
        for r in self.get_rows():
            if self._norm(r["Approval"]) == "NO" and self._norm(r["Apply_Status"]) == "WAITING":
                self.update_row(r["Job_ID"], Apply_Status="Skipped")
                log.info("Skipped (NO): %s - %s", r["Company"], r["Role"])
                n += 1
        return n
