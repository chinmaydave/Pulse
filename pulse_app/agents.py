from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from threading import Event, Lock, Thread
from time import monotonic

from .email_service import EmailService
from .excel_repository import ExcelRepository
from .models import ExpirationTarget, ReminderMessage


class ReminderAgent:
    def __init__(self, repository: ExcelRepository, mailer: EmailService, app_base_url: str):
        self.repository = repository
        self.mailer = mailer
        self.app_base_url = app_base_url.rstrip("/")

    def build_message(self, target: ExpirationTarget) -> ReminderMessage:
        escalate = target.is_overdue
        subject_prefix = "Escalation" if escalate else "Reminder"
        subject = f"{subject_prefix}: {target.field_name} expiring for {target.employee_name}"
        body = (
            f"Employee: {target.employee_name}\n"
            f"Title: {target.employee_id}\n"
            f"Field: {target.field_name}\n"
            f"Expiration date: {target.expiration_date:%Y-%m-%d}\n"
            f"Days until due: {target.days_until_due}\n"
            f"Status: {self._status_text(target)}\n\n"
            f"Please review the expiring item and take action accordingly."
        )
        return ReminderMessage(
            target_key=target.target_key,
            recipient=target.email,
            subject=subject,
            body=body,
            sender=target.manager_email,
            escalate=escalate,
        )

    def _status_text(self, target: ExpirationTarget) -> str:
        if target.is_overdue:
            return "Overdue"
        if target.days_until_due <= 7:
            return "Due within 7 days"
        if target.days_until_due <= 14:
            return "Due within 14 days"
        if target.days_until_due <= 30:
            return "Due within 30 days"
        return "Upcoming"

    def send_for_target(self, target: ExpirationTarget) -> dict[str, str]:
        last_sent = self.repository.last_sent_for_target(target)
        if last_sent:
            sent_at = last_sent["timestamp"]
            return {
                "target_key": target.target_key,
                "employee_id": target.employee_id,
                "recipient": target.email,
                "sender": target.manager_email,
                "status": "skipped",
                "detail": f"Already sent at {sent_at}.",
            }

        message = self.build_message(target)
        result = self.mailer.send(message)
        self.repository.record_reminder(
            target=target,
            recipient=message.recipient,
            subject=message.subject,
            message_preview=message.body,
            channel=result.channel,
            status=result.status,
        )
        return {
            "target_key": target.target_key,
            "employee_id": target.employee_id,
            "recipient": message.recipient,
            "sender": message.sender,
            "status": result.status,
            "detail": result.detail,
        }

    def send_pending(self, days_ahead: int, expiration_filter: str | None = None) -> list[dict[str, str]]:
        return [
            self.send_for_target(target)
            for target in self.repository.pending_for_reminder(days_ahead, expiration_filter)
        ]

    def due_for_automatic_send(
        self,
        days_ahead: int,
        cooldown_hours: int,
        pending_targets: list[ExpirationTarget] | None = None,
    ) -> list[ExpirationTarget]:
        cutoff = datetime.now() - timedelta(hours=cooldown_hours)
        candidates = pending_targets if pending_targets is not None else self.repository.pending_for_reminder(days_ahead)
        return [
            target
            for target in candidates
            if self.repository.last_reminder_for_target(target.target_key) is None
            or self.repository.last_reminder_for_target(target.target_key) <= cutoff
        ]

    def send_due_automatic(self, days_ahead: int, cooldown_hours: int) -> list[dict[str, str]]:
        return [
            self.send_for_target(target)
            for target in self.due_for_automatic_send(days_ahead, cooldown_hours)
        ]


@dataclass(frozen=True)
class AgentSnapshot:
    enabled: bool
    running: bool
    email_backend: str
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
    email_backend: str
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
        targets = [
            target
            for target in agent.due_for_automatic_send(self.days_ahead, self.cooldown_hours)
            if self._not_attempted_recently(target)
        ]

        results = []
        for target in targets:
            self._remember_attempt(target)
            results.append(agent.send_for_target(target))

        with self._lock:
            self._processed_count += len(results)
            self._last_result = f"Processed {len(results)} reminder(s)."
        return results

    def snapshot(self) -> AgentSnapshot:
        with self._lock:
            return AgentSnapshot(
                enabled=self.enabled,
                running=bool(self._thread and self._thread.is_alive()),
                email_backend=self.email_backend,
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

    def _not_attempted_recently(self, target: ExpirationTarget) -> bool:
        attempted_at = self._recent_attempts.get(target.target_key)
        if attempted_at is None:
            return True
        cooldown_seconds = self.cooldown_hours * 60 * 60
        return monotonic() - attempted_at >= cooldown_seconds

    def _remember_attempt(self, target: ExpirationTarget) -> None:
        self._recent_attempts[target.target_key] = monotonic()

    def _set_result(self, result: str) -> None:
        with self._lock:
            self._last_result = result
