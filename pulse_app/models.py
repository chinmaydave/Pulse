from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime


STATUS_OPEN = "Open"
STATUS_IN_PROGRESS = "In Progress"
STATUS_SUBMITTED = "Submitted"
STATUS_COMPLETED = "Completed"
STATUS_ESCALATED = "Escalated"

ACTIVE_STATUSES = {STATUS_OPEN, STATUS_IN_PROGRESS, STATUS_ESCALATED}


@dataclass(frozen=True)
class RequestRecord:
    request_id: str
    title: str
    category: str
    assigned_to: str
    associate_email: str
    manager_email: str
    due_date: date
    status: str
    priority: str
    source_system: str
    requested_change: str
    current_value: str
    submitted_value: str
    notes: str
    created_at: datetime
    submitted_at: datetime | None
    last_reminder_at: datetime | None
    reminder_count: int
    escalated_at: datetime | None

    @property
    def is_pending(self) -> bool:
        return self.status in ACTIVE_STATUSES

    @property
    def is_overdue(self) -> bool:
        return self.is_pending and self.due_date < date.today()

    @property
    def days_until_due(self) -> int:
        return (self.due_date - date.today()).days


@dataclass(frozen=True)
class EmployeeRecord:
    title: str
    name: str
    email: str
    manager_email: str
    passport_valid_until: date | None

    @property
    def employee_id(self) -> str:
        return self.title

    def expiration_fields(self) -> list[tuple[str, date | None]]:
        return [
            ("Passport Expiry Date", self.passport_valid_until),
        ]

    def expiration_targets(self) -> list["ExpirationTarget"]:
        today = date.today()
        targets: list[ExpirationTarget] = []
        for field_name, expiration_date in self.expiration_fields():
            if expiration_date is None:
                continue
            days_until_due = (expiration_date - today).days
            status = "overdue" if expiration_date < today else "due_soon" if days_until_due <= 0 else "upcoming"
            targets.append(
                ExpirationTarget(
                    employee_id=self.employee_id,
                    employee_name=self.name,
                    email=self.email,
                    manager_email=self.manager_email,
                    field_name=field_name,
                    expiration_date=expiration_date,
                    days_until_due=days_until_due,
                    status=status,
                )
            )
        return targets


@dataclass(frozen=True)
class ExpirationTarget:
    employee_id: str
    employee_name: str
    email: str
    manager_email: str
    field_name: str
    expiration_date: date
    days_until_due: int
    status: str

    @property
    def is_overdue(self) -> bool:
        return self.expiration_date < date.today()

    @property
    def target_key(self) -> str:
        return f"{self.employee_id}|{self.field_name}"


@dataclass(frozen=True)
class ReminderMessage:
    target_key: str
    recipient: str
    subject: str
    body: str
    escalate: bool = False
