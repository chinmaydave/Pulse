from dataclasses import dataclass
from pathlib import Path
import os


BASE_DIR = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class AppConfig:
    workbook_path: Path
    upload_dir: Path = BASE_DIR / "data" / "uploads"
    app_base_url: str = "http://127.0.0.1:5000"
    use_outlook: bool = False
    auto_reminders_enabled: bool = False
    reminder_days_ahead: int = 3
    reminder_scan_interval_seconds: int = 300
    reminder_cooldown_hours: int = 24
    host: str = "0.0.0.0"
    port: int = 5000
    debug: bool = False

    @classmethod
    def from_env(cls) -> "AppConfig":
        workbook_path = Path(
            os.getenv("PULSE_WORKBOOK_PATH", BASE_DIR / "data" / "pulse_requests_mock.xlsx")
        )
        return cls(
            workbook_path=workbook_path,
            upload_dir=Path(os.getenv("PULSE_UPLOAD_DIR", BASE_DIR / "data" / "uploads")),
            app_base_url=os.getenv("PULSE_APP_BASE_URL", "http://127.0.0.1:5000"),
            use_outlook=os.getenv("PULSE_USE_OUTLOOK", "false").lower() == "true",
            auto_reminders_enabled=os.getenv("PULSE_AUTO_REMINDERS", "false").lower() == "true",
            reminder_days_ahead=int(os.getenv("PULSE_REMINDER_DAYS_AHEAD", "3")),
            reminder_scan_interval_seconds=int(os.getenv("PULSE_REMINDER_SCAN_INTERVAL_SECONDS", "300")),
            reminder_cooldown_hours=int(os.getenv("PULSE_REMINDER_COOLDOWN_HOURS", "24")),
            host=os.getenv("PULSE_HOST", "0.0.0.0"),
            port=int(os.getenv("PULSE_PORT", "5000")),
            debug=os.getenv("PULSE_DEBUG", "false").lower() == "true",
        )
