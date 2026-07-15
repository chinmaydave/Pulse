from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pulse_app.excel_repository import ExcelRepository  # noqa: E402


def main() -> None:
    workbook_path = ROOT / "data" / "pulse_requests_mock.xlsx"
    if workbook_path.exists():
        workbook_path.unlink()
    ExcelRepository(workbook_path)
    print(f"Created {workbook_path}")


if __name__ == "__main__":
    main()
