from __future__ import annotations

import base64
import json
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


GRAPH_SCOPE = "offline_access User.Read Files.ReadWrite"
DEVICE_CODE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"


class GraphAuthPending(Exception):
    pass


def start_device_flow(client_id: str, tenant: str = "common") -> dict:
    client_id = client_id.strip()
    tenant = (tenant or "common").strip()
    if not client_id:
        raise ValueError("Enter a Microsoft app client ID for private OneDrive sign-in.")

    return _post_json(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode",
        {
            "client_id": client_id,
            "scope": GRAPH_SCOPE,
        },
    )


def complete_device_flow(client_id: str, tenant: str, device_code: str) -> dict:
    token = _post_json(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        {
            "client_id": client_id,
            "grant_type": DEVICE_CODE_GRANT,
            "device_code": device_code,
        },
        pending_is_ok=True,
    )
    return _with_expiry(token)


def refresh_token(client_id: str, tenant: str, refresh_token_value: str) -> dict:
    token = _post_json(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token_value,
            "scope": GRAPH_SCOPE,
        },
    )
    return _with_expiry(token)


def download_private_workbook(source_url: str, destination: Path, token_cache: dict) -> dict:
    access_token, updated_cache = access_token_from_cache(token_cache)
    request = Request(
        f"https://graph.microsoft.com/v1.0/shares/{_share_token(source_url)}/driveItem/content",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with urlopen(request, timeout=45) as response:
            content = response.read()
    except HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise ValueError(f"Microsoft Graph could not download that private workbook. {details}") from exc
    except (OSError, URLError) as exc:
        raise ValueError(f"Microsoft Graph could not be reached. {exc}") from exc

    if not content.startswith(b"PK"):
        raise ValueError("Microsoft Graph returned a response, but it was not an Excel workbook.")
    destination.write_bytes(content)
    return updated_cache


def access_token_from_cache(token_cache: dict) -> tuple[str, dict]:
    if not token_cache:
        raise ValueError("Private OneDrive is not signed in yet.")
    if token_cache.get("access_token") and float(token_cache.get("expires_at", 0)) > time.time() + 60:
        return token_cache["access_token"], token_cache
    if not token_cache.get("refresh_token"):
        raise ValueError("Private OneDrive sign-in expired. Start Microsoft sign-in again.")

    updated = refresh_token(
        token_cache.get("client_id", ""),
        token_cache.get("tenant", "common"),
        token_cache["refresh_token"],
    )
    updated["client_id"] = token_cache.get("client_id", "")
    updated["tenant"] = token_cache.get("tenant", "common")
    return updated["access_token"], updated


def _post_json(url: str, form: dict[str, str], pending_is_ok: bool = False) -> dict:
    request = Request(
        url,
        data=urlencode(form).encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="ignore")
        try:
            payload = json.loads(body)
        except json.JSONDecodeError:
            payload = {"error_description": body}
        if pending_is_ok and payload.get("error") == "authorization_pending":
            raise GraphAuthPending("Microsoft sign-in is not complete yet.")
        error = payload.get("error_description") or payload.get("error") or body
        raise ValueError(error) from exc
    except (OSError, URLError) as exc:
        raise ValueError(f"Microsoft sign-in could not be reached. {exc}") from exc


def _with_expiry(token: dict) -> dict:
    token["expires_at"] = time.time() + int(token.get("expires_in", 3600))
    return token


def _share_token(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.strip().encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{encoded}"
