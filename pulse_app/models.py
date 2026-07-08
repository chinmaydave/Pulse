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
class ReminderMessage:
    request_id: str
    recipient: str
    subject: str
    body: str
    escalate: bool = False
