from __future__ import annotations

from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen


def download_onedrive_workbook(source_url: str, destination: Path) -> None:
    url = source_url.strip()
    if not url:
        raise ValueError("Enter a OneDrive Excel URL.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    last_error: Exception | None = None
    for candidate in _download_candidates(url):
        try:
            _download(candidate, destination)
            return
        except Exception as exc:  # noqa: BLE001 - try the next URL shape before failing
            last_error = exc

    raise ValueError(
        "Could not download an Excel workbook from that OneDrive URL. "
        "Use a direct download link or a sharing link the app can access."
    ) from last_error


def _download_candidates(url: str) -> list[str]:
    return [url, _with_query_value(url, "download", "1")]


def _with_query_value(url: str, key: str, value: str) -> str:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query[key] = value
    return urlunparse(parsed._replace(query=urlencode(query)))


def _download(url: str, destination: Path) -> None:
    request = Request(url, headers={"User-Agent": "Pulse/1.0"})
    with urlopen(request, timeout=30) as response:
        content = response.read()

    if not content.startswith(b"PK"):
        raise ValueError("Downloaded content is not an .xlsx workbook.")

    destination.write_bytes(content)
