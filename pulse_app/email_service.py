from __future__ import annotations

from dataclasses import dataclass
from sys import platform
from typing import TYPE_CHECKING

from .models import ReminderMessage
from .sender_credentials import is_gmail_address, load_sender_credentials

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


class GmailEmailService(EmailService):
    def __init__(self, config: "AppConfig"):
        self.config = config

    def send(self, message: ReminderMessage) -> SendResult:
        sender = message.sender.strip().lower()
        password = load_sender_credentials(self.config).get(sender)
        if not password:
            return SendResult(
                channel="gmail",
                status="error",
                detail=f"Gmail app password is not saved for {sender}.",
            )

        import smtplib
        from email.message import EmailMessage

        email = EmailMessage()
        email["From"] = sender
        email["To"] = message.recipient
        email["Subject"] = message.subject
        email.set_content(message.body)

        try:
            with smtplib.SMTP("smtp.gmail.com", 587, timeout=15) as server:
                server.starttls()
                server.login(sender, password)
                server.send_message(email)
        except Exception as exc:  # noqa: BLE001 - surface provider errors in the UI/log
            return SendResult(
                channel="gmail",
                status="error",
                detail=f"Gmail send failed for {sender}: {exc}",
            )

        return SendResult(
            channel="gmail",
            status="sent",
            detail=f"Sent through Gmail from {sender} to {message.recipient}.",
        )


class ExcelRoutedEmailService(EmailService):
    def __init__(self, config: "AppConfig"):
        self.gmail = GmailEmailService(config)
        self.outlook = OutlookEmailService()

    def send(self, message: ReminderMessage) -> SendResult:
        if message.sender and is_gmail_address(message.sender):
            return self.gmail.send(message)
        return self.outlook.send(message)


def email_service(config: "AppConfig | bool") -> EmailService:
    if isinstance(config, bool):
        return OutlookEmailService() if config else DevelopmentEmailService()

    backend = getattr(config, "email_backend", "dev")
    if backend == "excel":
        return ExcelRoutedEmailService(config)
    if backend == "outlook":
        return OutlookEmailService()
    return DevelopmentEmailService()
