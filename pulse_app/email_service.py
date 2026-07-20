from __future__ import annotations

from dataclasses import dataclass
from sys import platform
from typing import TYPE_CHECKING

from .models import ReminderMessage

if TYPE_CHECKING:
    from .config import AppConfig, SmtpSettings


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


class SmtpEmailService(EmailService):
    """Sends reminders over SMTP.

    Works on any OS (unlike the Outlook COM path), which makes it the backend to
    use for local end-to-end testing against a catcher such as MailHog or Mailtrap,
    or a real provider via app-password credentials.
    """

    def __init__(self, settings: "SmtpSettings"):
        self.settings = settings

    def send(self, message: ReminderMessage) -> SendResult:
        import smtplib
        from email.message import EmailMessage

        email = EmailMessage()
        sender = message.sender or self.settings.from_address
        email["From"] = sender
        email["To"] = message.recipient
        email["Subject"] = message.subject
        email.set_content(message.body)

        try:
            if self.settings.use_ssl:
                server = smtplib.SMTP_SSL(
                    self.settings.host, self.settings.port, timeout=self.settings.timeout
                )
            else:
                server = smtplib.SMTP(
                    self.settings.host, self.settings.port, timeout=self.settings.timeout
                )
            with server:
                if self.settings.use_tls:
                    server.starttls()
                if self.settings.username:
                    server.login(self.settings.username, self.settings.password)
                server.send_message(email)
        except Exception as exc:  # noqa: BLE001 - surface any send failure to the log
            return SendResult(
                channel="smtp",
                status="error",
                detail=f"SMTP send failed: {exc}",
            )

        return SendResult(
            channel="smtp",
            status="sent",
            detail=f"Sent through SMTP from {sender} to {message.recipient}.",
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
    """Select an email backend.

    Accepts an AppConfig (preferred) and reads ``email_backend`` / ``smtp`` from it.
    A plain bool is still accepted for backward compatibility with the original
    ``email_service(use_outlook)`` signature.
    """
    if isinstance(config, bool):
        return OutlookEmailService() if config else DevelopmentEmailService()

    backend = getattr(config, "email_backend", "dev")
    if backend == "outlook":
        return OutlookEmailService()
    if backend == "smtp":
        return SmtpEmailService(config.smtp)
    return DevelopmentEmailService()
