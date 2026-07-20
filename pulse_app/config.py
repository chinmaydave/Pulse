from dataclasses import dataclass, field
from pathlib import Path
import os


BASE_DIR = Path(__file__).resolve().parent.parent


def _env_bool(name: str, default: bool = False) -> bool:
    return os.getenv(name, str(default)).lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class SmtpSettings:
    host: str = "localhost"
    port: int = 1025
    username: str = ""
    password: str = ""
    use_tls: bool = False
    use_ssl: bool = False
    from_address: str = "pulse@localhost"
    timeout: int = 10

    @classmethod
    def from_env(cls) -> "SmtpSettings":
        return cls(
            host=os.getenv("PULSE_SMTP_HOST", "localhost"),
            port=int(os.getenv("PULSE_SMTP_PORT", "1025")),
            username=os.getenv("PULSE_SMTP_USER", ""),
            password=os.getenv("PULSE_SMTP_PASSWORD", ""),
            use_tls=_env_bool("PULSE_SMTP_USE_TLS", False),
            use_ssl=_env_bool("PULSE_SMTP_USE_SSL", False),
            from_address=os.getenv("PULSE_SMTP_FROM", "pulse@localhost"),
            timeout=int(os.getenv("PULSE_SMTP_TIMEOUT", "10")),
        )


@dataclass(frozen=True)
class OneDriveSettings:
    file_id: str = ""
    drive_id: str = ""

    @classmethod
    def from_env(cls) -> "OneDriveSettings":
        return cls(
            file_id=os.getenv("PULSE_ONEDRIVE_FILE_ID", ""),
            drive_id=os.getenv("PULSE_ONEDRIVE_DRIVE_ID", ""),
        )


def _resolve_backend() -> str:
    """dev (default) | smtp | outlook.

    PULSE_EMAIL_BACKEND wins if set. Otherwise PULSE_USE_OUTLOOK=true is honored
    for backward compatibility with the original Outlook-only flag.
    """
    backend = os.getenv("PULSE_EMAIL_BACKEND", "").strip().lower()
    if backend:
        return backend
    if _env_bool("PULSE_USE_OUTLOOK", False):
        return "outlook"
    return "dev"


@dataclass(frozen=True)
class AppConfig:
    workbook_path: Path
    upload_dir: Path = BASE_DIR / "data" / "uploads"
    app_base_url: str = "http://127.0.0.1:5000"
    auto_reminders_enabled: bool = False
    email_backend: str = "dev"
    smtp: SmtpSettings = field(default_factory=SmtpSettings)
    data_backend: str = "excel"
    onedrive: OneDriveSettings = field(default_factory=OneDriveSettings)
    reminder_days_ahead: int = 3
    reminder_scan_interval_seconds: int = 300
    reminder_cooldown_hours: int = 24
    host: str = "0.0.0.0"
    port: int = 5000
    debug: bool = False

    @property
    def use_outlook(self) -> bool:
        # Preserved so existing callers/templates that read use_outlook still work.
        return self.email_backend == "outlook"

    @classmethod
    def from_env(cls) -> "AppConfig":
        workbook_path = Path(
            os.getenv("PULSE_WORKBOOK_PATH", BASE_DIR / "data" / "pulse_passport_expirations_mock.xlsx")
        )
        return cls(
            workbook_path=workbook_path,
            upload_dir=Path(os.getenv("PULSE_UPLOAD_DIR", BASE_DIR / "data" / "uploads")),
            app_base_url=os.getenv("PULSE_APP_BASE_URL", "http://127.0.0.1:5000"),
            auto_reminders_enabled=os.getenv("PULSE_AUTO_REMINDERS", "false").lower() == "true",
            email_backend=_resolve_backend(),
            smtp=SmtpSettings.from_env(),
            data_backend=os.getenv("PULSE_DATA_BACKEND", "excel").strip().lower(),
            onedrive=OneDriveSettings.from_env(),
            reminder_days_ahead=int(os.getenv("PULSE_REMINDER_DAYS_AHEAD", "3")),
            reminder_scan_interval_seconds=int(os.getenv("PULSE_REMINDER_SCAN_INTERVAL_SECONDS", "300")),
            reminder_cooldown_hours=int(os.getenv("PULSE_REMINDER_COOLDOWN_HOURS", "24")),
            host=os.getenv("PULSE_HOST", "0.0.0.0"),
            port=int(os.getenv("PULSE_PORT", "5000")),
            debug=os.getenv("PULSE_DEBUG", "false").lower() == "true",
        )
