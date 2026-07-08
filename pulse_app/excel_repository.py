from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from threading import Lock
from typing import Any

from openpyxl import Workbook, load_workbook

from .models import (
    ACTIVE_STATUSES,
    RequestRecord,
    STATUS_COMPLETED,
    STATUS_ESCALATED,
    STATUS_IN_PROGRESS,
    STATUS_SUBMITTED,
)


REQUEST_HEADERS = [
    "request_id",
    "title",
    "category",
    "assigned_to",
    "associate_email",
    "manager_email",
    "due_date",
    "status",
    "priority",
    "source_system",
    "requested_change",
    "current_value",
    "submitted_value",
    "notes",
    "created_at",
    "submitted_at",
    "last_reminder_at",
    "reminder_count",
    "escalated_at",
]

AUDIT_HEADERS = ["timestamp", "request_id", "actor", "action", "details"]
REMINDER_HEADERS = [
    "timestamp",
    "request_id",
    "recipient",
    "subject",
    "channel",
    "status",
    "message_preview",
]


class ExcelRepository:
    def __init__(self, workbook_path: Path):
        self.workbook_path = Path(workbook_path)
        self._lock = Lock()
        if not self.workbook_path.exists():
            self.create_mock_workbook()

    def create_mock_workbook(self) -> None:
        self.workbook_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = Workbook()
        requests = workbook.active
        requests.title = "Requests"
        requests.append(REQUEST_HEADERS)

        today = date.today()
        rows = [
            [
                "REQ-1001",
                "Update emergency contact",
                "Employee Data",
                "Avery Johnson",
                "avery.johnson@example.com",
                "manager.hr@example.com",
                today,
                "Open",
                "High",
                "HRIS",
                "Confirm and update emergency contact details.",
                "Primary contact: Morgan Johnson, 555-0101",
                "",
                "",
                datetime.now(),
                None,
                None,
                0,
                None,
            ],
            [
                "REQ-1002",
                "Validate cost center",
                "Finance",
                "Blake Smith",
                "blake.smith@example.com",
                "manager.finance@example.com",
                today + timedelta(days=2),
                "In Progress",
                "Medium",
                "ERP",
                "Validate the employee cost center and submit corrections.",
                "CC-4820 Operations",
                "",
                "Associate opened request.",
                datetime.now(),
                None,
                None,
                1,
                None,
            ],
            [
                "REQ-1003",
                "Submit license renewal details",
                "Compliance",
                "Casey Lee",
                "casey.lee@example.com",
                "manager.compliance@example.com",
                today - timedelta(days=1),
                "Open",
                "Critical",
                "Compliance Tracker",
                "Provide renewed license number and expiration date.",
                "License expires this month.",
                "",
                "",
                datetime.now(),
                None,
                None,
                2,
                None,
            ],
            [
                "REQ-1004",
                "Confirm team directory data",
                "Operations",
                "Devon Patel",
                "devon.patel@example.com",
                "manager.ops@example.com",
                today + timedelta(days=5),
                "Submitted",
                "Low",
                "Directory",
                "Confirm title, location, and desk phone.",
                "Analyst, Chicago, 555-0144",
                "Analyst, Chicago, 555-0188",
                "Desk phone changed.",
                datetime.now(),
                datetime.now(),
                None,
                0,
                None,
            ],
        ]

        for row in rows:
            requests.append(row)

        audit = workbook.create_sheet("AuditLog")
        audit.append(AUDIT_HEADERS)
        audit.append([datetime.now(), "SYSTEM", "seed", "created", "Mock workbook initialized."])

        reminders = workbook.create_sheet("ReminderLog")
        reminders.append(REMINDER_HEADERS)

        workbook.save(self.workbook_path)

    def list_requests(self) -> list[RequestRecord]:
        with self._lock:
            workbook = load_workbook(self.workbook_path)
            sheet = workbook["Requests"]
            return [self._row_to_record(row) for row in self._iter_rows(sheet)]

    def get_request(self, request_id: str) -> RequestRecord | None:
        return next((record for record in self.list_requests() if record.request_id == request_id), None)

    def pending_for_reminder(self, days_ahead: int) -> list[RequestRecord]:
        today = date.today()
        horizon = today.toordinal() + days_ahead
        return [
            record
            for record in self.list_requests()
            if record.status in ACTIVE_STATUSES and record.due_date.toordinal() <= horizon
        ]

    def submit_response(self, request_id: str, submitted_value: str, notes: str, actor: str) -> None:
        with self._lock:
            workbook = load_workbook(self.workbook_path)
            sheet = workbook["Requests"]
            row_idx = self._find_row(sheet, request_id)
            if row_idx is None:
                raise KeyError(f"Request {request_id} not found")

            self._set(sheet, row_idx, "submitted_value", submitted_value)
            self._set(sheet, row_idx, "notes", notes)
            self._set(sheet, row_idx, "status", STATUS_SUBMITTED)
            self._set(sheet, row_idx, "submitted_at", datetime.now())
            self._append_audit(workbook, request_id, actor, "submitted", "Associate submitted form response.")
            workbook.save(self.workbook_path)

    def update_status(self, request_id: str, status: str, actor: str, details: str = "") -> None:
        valid_statuses = {STATUS_IN_PROGRESS, STATUS_SUBMITTED, STATUS_COMPLETED, STATUS_ESCALATED}
        if status not in valid_statuses:
            raise ValueError(f"Unsupported status: {status}")

        with self._lock:
            workbook = load_workbook(self.workbook_path)
            sheet = workbook["Requests"]
            row_idx = self._find_row(sheet, request_id)
            if row_idx is None:
                raise KeyError(f"Request {request_id} not found")
            self._set(sheet, row_idx, "status", status)
            if status == STATUS_ESCALATED:
                self._set(sheet, row_idx, "escalated_at", datetime.now())
            self._append_audit(workbook, request_id, actor, "status_updated", details or f"Status changed to {status}.")
            workbook.save(self.workbook_path)

    def record_reminder(
        self,
        request_id: str,
        recipient: str,
        subject: str,
        message_preview: str,
        channel: str,
        status: str,
        escalate: bool,
    ) -> None:
        with self._lock:
            workbook = load_workbook(self.workbook_path)
            sheet = workbook["Requests"]
            row_idx = self._find_row(sheet, request_id)
            if row_idx is None:
                raise KeyError(f"Request {request_id} not found")

            now = datetime.now()
            current_count = self._get(sheet, row_idx, "reminder_count") or 0
            self._set(sheet, row_idx, "last_reminder_at", now)
            self._set(sheet, row_idx, "reminder_count", int(current_count) + 1)
            if escalate:
                self._set(sheet, row_idx, "status", STATUS_ESCALATED)
                self._set(sheet, row_idx, "escalated_at", now)

            reminders = workbook["ReminderLog"]
            reminders.append([now, request_id, recipient, subject, channel, status, message_preview[:180]])
            self._append_audit(
                workbook,
                request_id,
                "system",
                "reminder_sent",
                f"{channel} reminder recorded for {recipient}.",
            )
            workbook.save(self.workbook_path)

    def audit_rows(self) -> list[dict[str, Any]]:
        return self._sheet_rows("AuditLog")

    def reminder_rows(self) -> list[dict[str, Any]]:
        return self._sheet_rows("ReminderLog")

    def summary(self) -> dict[str, int]:
        records = self.list_requests()
        return {
            "total": len(records),
            "open": sum(1 for record in records if record.status == "Open"),
            "in_progress": sum(1 for record in records if record.status == STATUS_IN_PROGRESS),
            "submitted": sum(1 for record in records if record.status == STATUS_SUBMITTED),
            "completed": sum(1 for record in records if record.status == STATUS_COMPLETED),
            "escalated": sum(1 for record in records if record.status == STATUS_ESCALATED),
            "overdue": sum(1 for record in records if record.is_overdue),
        }

    def _sheet_rows(self, sheet_name: str) -> list[dict[str, Any]]:
        with self._lock:
            workbook = load_workbook(self.workbook_path)
            sheet = workbook[sheet_name]
            headers = [cell.value for cell in sheet[1]]
            rows = []
            for raw_row in sheet.iter_rows(min_row=2, values_only=True):
                if any(value is not None for value in raw_row):
                    rows.append(dict(zip(headers, raw_row)))
            return rows

    def _iter_rows(self, sheet: Any) -> list[dict[str, Any]]:
        headers = [cell.value for cell in sheet[1]]
        return [
            dict(zip(headers, raw_row))
            for raw_row in sheet.iter_rows(min_row=2, values_only=True)
            if any(value is not None for value in raw_row)
        ]

    def _row_to_record(self, row: dict[str, Any]) -> RequestRecord:
        due_date = row["due_date"]
        if isinstance(due_date, datetime):
            due_date = due_date.date()
        return RequestRecord(
            request_id=row["request_id"],
            title=row["title"],
            category=row["category"],
            assigned_to=row["assigned_to"],
            associate_email=row["associate_email"],
            manager_email=row["manager_email"],
            due_date=due_date,
            status=row["status"],
            priority=row["priority"],
            source_system=row["source_system"],
            requested_change=row["requested_change"],
            current_value=row["current_value"] or "",
            submitted_value=row["submitted_value"] or "",
            notes=row["notes"] or "",
            created_at=row["created_at"],
            submitted_at=row["submitted_at"],
            last_reminder_at=row["last_reminder_at"],
            reminder_count=int(row["reminder_count"] or 0),
            escalated_at=row["escalated_at"],
        )

    def _find_row(self, sheet: Any, request_id: str) -> int | None:
        request_col = REQUEST_HEADERS.index("request_id") + 1
        for row_idx in range(2, sheet.max_row + 1):
            if sheet.cell(row=row_idx, column=request_col).value == request_id:
                return row_idx
        return None

    def _set(self, sheet: Any, row_idx: int, field: str, value: Any) -> None:
        sheet.cell(row=row_idx, column=REQUEST_HEADERS.index(field) + 1, value=value)

    def _get(self, sheet: Any, row_idx: int, field: str) -> Any:
        return sheet.cell(row=row_idx, column=REQUEST_HEADERS.index(field) + 1).value

    def _append_audit(self, workbook: Any, request_id: str, actor: str, action: str, details: str) -> None:
        workbook["AuditLog"].append([datetime.now(), request_id, actor, action, details])
