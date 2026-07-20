from __future__ import annotations

import base64
import json
import re
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from urllib.error import HTTPError, URLError
from urllib.request import HTTPCookieProcessor, HTTPRedirectHandler, Request, build_opener


REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    ),
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/html,*/*",
}


def download_onedrive_workbook(source_url: str, destination: Path) -> None:
    url = source_url.strip()
    if not url:
        raise ValueError("Enter a OneDrive Excel URL.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    last_error: Exception | None = None
    candidates = _download_candidates(url)
    seen: set[str] = set()
    while candidates:
        candidate = candidates.pop(0)
        if candidate in seen:
            continue
        seen.add(candidate)
        try:
            candidates.extend(_download(candidate, destination))
            if destination.exists():
                return
        except Exception as exc:  # noqa: BLE001 - try the next URL shape before failing
            last_error = exc

    raise ValueError(
        "Could not download an Excel workbook from that OneDrive link. "
        "Confirm the link is a share link with view access. Private or sign-in-only "
        "links require Microsoft Graph authentication."
    ) from last_error


def _download_candidates(url: str) -> list[str]:
    candidates = [
        url,
        _with_query_value(url, "download", "1"),
        _onedrive_api_share_content_url(url),
        _graph_share_content_url(url),
    ]
    redirect_url = _first_redirect_url(url)
    if redirect_url:
        candidates.insert(1, redirect_url)
        candidates.insert(2, _with_query_value(redirect_url, "download", "1"))

    direct_download = _onedrive_live_download_url(url)
    if direct_download:
        candidates.insert(1, direct_download)
    return [candidate for candidate in candidates if candidate]


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _first_redirect_url(url: str) -> str | None:
    opener = build_opener(_NoRedirect)
    try:
        opener.open(Request(url, headers=REQUEST_HEADERS), timeout=15)
    except HTTPError as exc:
        if 300 <= exc.code < 400:
            return exc.headers.get("Location")
    except (OSError, URLError):
        return None
    return None


def _share_token(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{encoded}"


def _onedrive_api_share_content_url(url: str) -> str:
    return f"https://api.onedrive.com/v1.0/shares/{_share_token(url)}/root/content"


def _graph_share_content_url(url: str) -> str:
    return f"https://graph.microsoft.com/v1.0/shares/{_share_token(url)}/driveItem/content"


def _onedrive_live_download_url(url: str) -> str | None:
    parsed = urlparse(url)
    if parsed.netloc.lower() != "onedrive.live.com":
        return None

    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    resid = query.get("resid")
    authkey = query.get("authkey")
    if not resid:
        return None

    download_query = {"resid": resid}
    if authkey:
        download_query["authkey"] = authkey
    return urlunparse(parsed._replace(path="/download", query=urlencode(download_query)))


def _with_query_value(url: str, key: str, value: str) -> str:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query[key] = value
    return urlunparse(parsed._replace(query=urlencode(query)))


def _download(url: str, destination: Path) -> list[str]:
    request = Request(url, headers=REQUEST_HEADERS)
    opener = build_opener(HTTPCookieProcessor(CookieJar()))
    with opener.open(request, timeout=30) as response:
        content = response.read()
        final_url = response.geturl()
        content_type = response.headers.get("content-type", "")

    if "login.live.com" in urlparse(final_url).netloc.lower():
        raise ValueError("OneDrive redirected the app to Microsoft sign-in instead of an Excel download.")
    if content.startswith(b"PK"):
        destination.write_bytes(content)
        return []
    if "json" in content_type:
        return _extract_download_candidates_from_json(content)
    return _extract_download_candidates_from_html(content, final_url)


def _extract_download_candidates_from_json(content: bytes) -> list[str]:
    try:
        data = json.loads(content.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return []

    candidates = []
    download_url = data.get("@microsoft.graph.downloadUrl") or data.get("@content.downloadUrl")
    web_url = data.get("webUrl")
    if download_url:
        candidates.append(download_url)
    if web_url:
        candidates.extend(_download_candidates(web_url))
    return candidates


def _extract_download_candidates_from_html(content: bytes, final_url: str) -> list[str]:
    text = content.decode("utf-8", errors="ignore")
    candidates = []
    for match in re.findall(r'"(?:downloadUrl|@content\.downloadUrl)"\s*:\s*"([^"]+)"', text):
        candidates.append(_json_unescape(match))
    for match in re.findall(r'https:\\/\\/[^"]+', text):
        candidate = _json_unescape(match)
        if "download" in candidate.lower() or ".xlsx" in candidate.lower():
            candidates.append(candidate)

    live_download = _onedrive_live_download_url(final_url)
    if live_download:
        candidates.append(live_download)
    return candidates


def _json_unescape(value: str) -> str:
    try:
        return json.loads(f'"{value}"')
    except json.JSONDecodeError:
        return value.replace("\\/", "/")
