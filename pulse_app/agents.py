from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from threading import Event, Lock, Thread
from time import monotonic

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

    def due_for_automatic_send(self, days_ahead: int, cooldown_hours: int) -> list[RequestRecord]:
        cutoff = datetime.now() - timedelta(hours=cooldown_hours)
        return [
            record
            for record in self.repository.pending_for_reminder(days_ahead)
            if record.last_reminder_at is None or record.last_reminder_at <= cutoff
        ]

    def send_due_automatic(self, days_ahead: int, cooldown_hours: int) -> list[dict[str, str]]:
        return [
            self.send_for_record(record)
            for record in self.due_for_automatic_send(days_ahead, cooldown_hours)
        ]


@dataclass(frozen=True)
class AgentSnapshot:
    enabled: bool
    running: bool
    use_outlook: bool
    days_ahead: int
    scan_interval_seconds: int
    cooldown_hours: int
    last_started_at: datetime | None = None
    last_scan_at: datetime | None = None
    last_result: str = "Not started."
    processed_count: int = 0
    error_count: int = 0


@dataclass
class AutomaticReminderAgent:
    repository: ExcelRepository
    mailer: EmailService
    app_base_url: str
    use_outlook: bool
    days_ahead: int
    scan_interval_seconds: int
    cooldown_hours: int
    enabled: bool
    _stop_event: Event = field(default_factory=Event, init=False)
    _thread: Thread | None = field(default=None, init=False)
    _lock: Lock = field(default_factory=Lock, init=False)
    _last_started_at: datetime | None = field(default=None, init=False)
    _last_scan_at: datetime | None = field(default=None, init=False)
    _last_result: str = field(default="Not started.", init=False)
    _processed_count: int = field(default=0, init=False)
    _error_count: int = field(default=0, init=False)
    _recent_attempts: dict[str, float] = field(default_factory=dict, init=False)

    def start(self) -> None:
        if not self.enabled:
            return
        if self._thread and self._thread.is_alive():
            return

        self._last_started_at = datetime.now()
        self._thread = Thread(target=self._run, name="pulse-automatic-reminders", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)

    def run_once(self) -> list[dict[str, str]]:
        if not self.enabled:
            self._set_result("Automatic reminders are disabled.")
            return []

        self._last_scan_at = datetime.now()
        agent = ReminderAgent(self.repository, self.mailer, self.app_base_url)
        records = [
            record
            for record in agent.due_for_automatic_send(self.days_ahead, self.cooldown_hours)
            if self._not_attempted_recently(record)
        ]

        results = []
        for record in records:
            self._remember_attempt(record)
            results.append(agent.send_for_record(record))

        with self._lock:
            self._processed_count += len(results)
            self._last_result = f"Processed {len(results)} reminder(s)."
        return results

    def snapshot(self) -> AgentSnapshot:
        with self._lock:
            return AgentSnapshot(
                enabled=self.enabled,
                running=bool(self._thread and self._thread.is_alive()),
                use_outlook=self.use_outlook,
                days_ahead=self.days_ahead,
                scan_interval_seconds=self.scan_interval_seconds,
                cooldown_hours=self.cooldown_hours,
                last_started_at=self._last_started_at,
                last_scan_at=self._last_scan_at,
                last_result=self._last_result,
                processed_count=self._processed_count,
                error_count=self._error_count,
            )

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                self.run_once()
            except Exception as exc:
                with self._lock:
                    self._error_count += 1
                    self._last_result = f"{type(exc).__name__}: {exc}"
            self._stop_event.wait(self.scan_interval_seconds)

    def _not_attempted_recently(self, record: RequestRecord) -> bool:
        attempted_at = self._recent_attempts.get(record.request_id)
        if attempted_at is None:
            return True
        cooldown_seconds = self.cooldown_hours * 60 * 60
        return monotonic() - attempted_at >= cooldown_seconds

    def _remember_attempt(self, record: RequestRecord) -> None:
        self._recent_attempts[record.request_id] = monotonic()

    def _set_result(self, result: str) -> None:
        with self._lock:
            self._last_result = result
