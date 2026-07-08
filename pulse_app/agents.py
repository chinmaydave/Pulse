from __future__ import annotations

from .email_service import EmailService
from .excel_repository import ExcelRepository
from .models import ReminderMessage, RequestRecord


class ReminderAgent:
    def __init__(self, repository: ExcelRepository, mailer: EmailService, app_base_url: str):
        self.repository = repository
        self.mailer = mailer
        self.app_base_url = app_base_url.rstrip("/")

    def build_message(self, record: RequestRecord) -> ReminderMessage:
        form_url = f"{self.app_base_url}/requests/{record.request_id}"
        escalate = record.is_overdue and record.reminder_count >= 2
        recipient = record.manager_email if escalate else record.associate_email
        subject_prefix = "Escalation" if escalate else "Reminder"
        subject = f"{subject_prefix}: {record.title} ({record.request_id})"
        body = (
            f"Request: {record.title}\n"
            f"Request ID: {record.request_id}\n"
            f"Assigned to: {record.assigned_to}\n"
            f"Due date: {record.due_date:%Y-%m-%d}\n"
            f"Status: {record.status}\n\n"
            f"Please complete the form here:\n{form_url}\n\n"
            f"Requested change:\n{record.requested_change}\n"
        )
        return ReminderMessage(
            request_id=record.request_id,
            recipient=recipient,
            subject=subject,
            body=body,
            escalate=escalate,
        )

    def send_for_record(self, record: RequestRecord) -> dict[str, str]:
        message = self.build_message(record)
        result = self.mailer.send(message)
        self.repository.record_reminder(
            request_id=record.request_id,
            recipient=message.recipient,
            subject=message.subject,
            message_preview=message.body,
            channel=result.channel,
            status=result.status,
            escalate=message.escalate,
        )
        return {
            "request_id": record.request_id,
            "recipient": message.recipient,
            "status": result.status,
            "detail": result.detail,
        }

    def send_pending(self, days_ahead: int) -> list[dict[str, str]]:
        return [self.send_for_record(record) for record in self.repository.pending_for_reminder(days_ahead)]
