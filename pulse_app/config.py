from dataclasses import dataclass, field
from pathlib import Path
import os


BASE_DIR = Path(__file__).resolve().parent.parent


def _env_bool(name: str, default: bool = False) -> bool:
    return os.getenv(name, str(default)).lower() in {"1", "true", "yes", "on"}


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


@dataclass(frozen=True)
class AppConfig:
    workbook_path: Path
    upload_dir: Path = BASE_DIR / "data" / "uploads"
    app_base_url: str = "http://127.0.0.1:5000"
    auto_reminders_enabled: bool = True
    email_backend: str = "outlook"
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
            os.getenv("PULSE_WORKBOOK_PATH", BASE_DIR / "data" / "EmployeeExpirations_OneDrive_Template.xlsx")
        )
        return cls(
            workbook_path=workbook_path,
            upload_dir=Path(os.getenv("PULSE_UPLOAD_DIR", BASE_DIR / "data" / "uploads")),
            app_base_url=os.getenv("PULSE_APP_BASE_URL", "http://127.0.0.1:5000"),
            auto_reminders_enabled=_env_bool("PULSE_AUTO_REMINDERS", True),
            email_backend="outlook",
            data_backend=os.getenv("PULSE_DATA_BACKEND", "excel").strip().lower(),
            onedrive=OneDriveSettings.from_env(),
            reminder_days_ahead=int(os.getenv("PULSE_REMINDER_DAYS_AHEAD", "3")),
            reminder_scan_interval_seconds=int(os.getenv("PULSE_REMINDER_SCAN_INTERVAL_SECONDS", "300")),
            reminder_cooldown_hours=int(os.getenv("PULSE_REMINDER_COOLDOWN_HOURS", "24")),
            host=os.getenv("PULSE_HOST", "0.0.0.0"),
            port=int(os.getenv("PULSE_PORT", "5000")),
            debug=os.getenv("PULSE_DEBUG", "false").lower() == "true",
        )
