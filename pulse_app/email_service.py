from __future__ import annotations

from dataclasses import dataclass
from sys import platform
from typing import TYPE_CHECKING

from .models import ReminderMessage

if TYPE_CHECKING:
    from .config import AppConfig


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
            detail=f"Development reminder prepared from {message.sender or 'default sender'} to {message.recipient}.",
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
        if message.sender:
            mail.SentOnBehalfOfName = message.sender
        mail.Subject = message.subject
        mail.Body = message.body
        mail.Send()
        sender_detail = f" from {message.sender}" if message.sender else ""
        return SendResult(
            channel="outlook",
            status="sent",
            detail=f"Sent through Outlook{sender_detail} to {message.recipient}.",
        )


def email_service(config: "AppConfig | bool") -> EmailService:
    if isinstance(config, bool):
        return OutlookEmailService() if config else DevelopmentEmailService()

    backend = getattr(config, "email_backend", "dev")
    if backend == "outlook":
        return OutlookEmailService()
    return DevelopmentEmailService()
