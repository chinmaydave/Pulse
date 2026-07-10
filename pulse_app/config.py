from dataclasses import dataclass
from pathlib import Path
import os


BASE_DIR = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class AppConfig:
    workbook_path: Path
    upload_dir: Path = BASE_DIR / "data" / "uploads"
    use_outlook: bool = False
    reminder_days_ahead: int = 3
    host: str = "0.0.0.0"
    port: int = 5000

    @classmethod
    def from_env(cls) -> "AppConfig":
        workbook_path = Path(
            os.getenv("PULSE_WORKBOOK_PATH", BASE_DIR / "data" / "pulse_requests_mock.xlsx")
        )
        return cls(
            workbook_path=workbook_path,
            upload_dir=Path(os.getenv("PULSE_UPLOAD_DIR", BASE_DIR / "data" / "uploads")),
            use_outlook=os.getenv("PULSE_USE_OUTLOOK", "false").lower() == "true",
            reminder_days_ahead=int(os.getenv("PULSE_REMINDER_DAYS_AHEAD", "3")),
            host=os.getenv("PULSE_HOST", "0.0.0.0"),
            port=int(os.getenv("PULSE_PORT", "5000")),
        )
