from dataclasses import dataclass
from pathlib import Path
import os


BASE_DIR = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class AppConfig:
    workbook_path: Path
    use_outlook: bool = False
    reminder_days_ahead: int = 3

    @classmethod
    def from_env(cls) -> "AppConfig":
        workbook_path = Path(
            os.getenv("PULSE_WORKBOOK_PATH", BASE_DIR / "data" / "pulse_requests_mock.xlsx")
        )
        return cls(
            workbook_path=workbook_path,
            use_outlook=os.getenv("PULSE_USE_OUTLOOK", "false").lower() == "true",
            reminder_days_ahead=int(os.getenv("PULSE_REMINDER_DAYS_AHEAD", "3")),
        )
