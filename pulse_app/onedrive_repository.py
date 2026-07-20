from __future__ import annotations

from pathlib import Path
from threading import Lock
from typing import Any, Callable

from .excel_repository import ExcelRepository
from .models import RequestRecord


class OneDriveRepository:
    """Wraps ExcelRepository to read/write an Excel file that lives on OneDrive.

    The actual Graph API calls are injected as callables so this class has no
    dependency on auth. Part 1a will supply real implementations; until then
    the stubs in __init__.py raise NotImplementedError on startup.

    Parameters
    ----------
    cache_path:
        Local path where the downloaded .xlsx is stored between requests.
    download_fn:
        Called as download_fn(dest_path: Path) -> None. Downloads the OneDrive
        file to dest_path. Implemented in Part 1a.
    upload_fn:
        Called as upload_fn(src_path: Path) -> None. Uploads the local file at
        src_path back to OneDrive. Implemented in Part 1a.
    """

    def __init__(
        self,
        cache_path: Path,
        download_fn: Callable[[Path], None],
        upload_fn: Callable[[Path], None],
    ) -> None:
        self._cache_path = Path(cache_path)
        self._download_fn = download_fn
        self._upload_fn = upload_fn
        self._sync_lock = Lock()

        self._download_fn(self._cache_path)
        self._inner = ExcelRepository(self._cache_path)

    # ------------------------------------------------------------------
    # workbook_path — keeps templates working without modification
    # ------------------------------------------------------------------

    @property
    def workbook_path(self) -> Path:
        return self._cache_path

    # ------------------------------------------------------------------
    # Read methods — delegate only, no upload needed
    # ------------------------------------------------------------------

    def list_requests(self) -> list[RequestRecord]:
        return self._inner.list_requests()

    def get_request(self, request_id: str) -> RequestRecord | None:
        return self._inner.get_request(request_id)

    def pending_for_reminder(self, days_ahead: int) -> list[RequestRecord]:
        return self._inner.pending_for_reminder(days_ahead)

    def audit_rows(self) -> list[dict[str, Any]]:
        return self._inner.audit_rows()

    def reminder_rows(self) -> list[dict[str, Any]]:
        return self._inner.reminder_rows()

    def summary(self) -> dict[str, int]:
        return self._inner.summary()

    # ------------------------------------------------------------------
    # Write methods — delegate then upload inside _sync_lock
    # ------------------------------------------------------------------

    def submit_response(
        self, request_id: str, submitted_value: str, notes: str, actor: str
    ) -> None:
        with self._sync_lock:
            self._inner.submit_response(request_id, submitted_value, notes, actor)
            self._upload_fn(self._cache_path)

    def update_status(
        self, request_id: str, status: str, actor: str, details: str = ""
    ) -> None:
        with self._sync_lock:
            self._inner.update_status(request_id, status, actor, details)
            self._upload_fn(self._cache_path)

    def record_reminder(
        self,
        request_id: str,
        recipient: str,
        subject: str,
        message_preview: str,
        channel: str,
        status: str,
        escalate: bool,
    ) -> None:
        with self._sync_lock:
            self._inner.record_reminder(
                request_id, recipient, subject, message_preview,
                channel, status, escalate,
            )
            self._upload_fn(self._cache_path)

    # ------------------------------------------------------------------
    # Workbook management
    # ------------------------------------------------------------------

    @staticmethod
    def validate_workbook(workbook_path: Path) -> None:
        ExcelRepository.validate_workbook(workbook_path)

    def ensure_workbook_ready(self) -> None:
        with self._sync_lock:
            self._inner.ensure_workbook_ready()
            self._upload_fn(self._cache_path)
