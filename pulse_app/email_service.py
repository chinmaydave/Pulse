from __future__ import annotations

from dataclasses import dataclass
from sys import platform

from .models import ReminderMessage


@dataclass(frozen=True)
class SendResult:
    channel: str
    status: str
    detail: str


class EmailService:
    def send(self, message: ReminderMessage) -> SendResult:
        raise NotImplementedError


class DevelopmentEmailService(EmailService):
    def send(self, message: ReminderMessage) -> SendResult:
        return SendResult(
            channel="development-log",
            status="logged",
            detail=f"Development reminder prepared for {message.recipient}.",
        )


class OutlookEmailService(EmailService):
    def send(self, message: ReminderMessage) -> SendResult:
        if platform != "win32":
            return SendResult(
                channel="outlook",
                status="skipped",
                detail="Outlook sending requires Windows with the Outlook desktop client.",
            )

        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = message.recipient
        mail.Subject = message.subject
        mail.Body = message.body
        mail.Send()
        return SendResult(channel="outlook", status="sent", detail="Sent through Outlook.")


def email_service(use_outlook: bool) -> EmailService:
    if use_outlook:
        return OutlookEmailService()
    return DevelopmentEmailService()
