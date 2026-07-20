from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path
from typing import Any

from .config import AppConfig, SmtpSettings


PROVIDERS = {
    "gmail": ("smtp.gmail.com", 587),
    "outlook": ("smtp-mail.outlook.com", 587),
    "microsoft365": ("smtp.office365.com", 587),
}


def email_settings_path(config: AppConfig) -> Path:
    return config.upload_dir / "email_settings.json"


def load_email_settings(config: AppConfig) -> dict[str, Any]:
    path = email_settings_path(config)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_email_settings(config: AppConfig, form: dict[str, str]) -> None:
    provider = form.get("provider", "gmail").strip().lower()
    default_host, default_port = PROVIDERS.get(provider, ("", 587))
    host = default_host if provider in PROVIDERS else form.get("host", "").strip()
    port = default_port if provider in PROVIDERS else int(form.get("port", "") or default_port)
    email_address = form.get("email_address", "").strip()
    username = form.get("username", "").strip() or email_address
    password = form.get("password", "")
    data = {
        "backend": "smtp",
        "provider": provider,
        "host": host,
        "port": port,
        "username": username,
        "password": password,
        "from_address": form.get("from_address", "").strip() or username,
        "use_tls": True,
        "use_ssl": False,
    }
    config.upload_dir.mkdir(parents=True, exist_ok=True)
    email_settings_path(config).write_text(json.dumps(data, indent=2), encoding="utf-8")


def clear_email_settings(config: AppConfig) -> None:
    email_settings_path(config).unlink(missing_ok=True)


def effective_email_config(config: AppConfig) -> AppConfig:
    settings = load_email_settings(config)
    if not settings:
        return config
    smtp = SmtpSettings(
        host=settings.get("host", "localhost"),
        port=int(settings.get("port", 1025)),
        username=settings.get("username", ""),
        password=settings.get("password", ""),
        use_tls=bool(settings.get("use_tls", True)),
        use_ssl=bool(settings.get("use_ssl", False)),
        from_address=settings.get("from_address", settings.get("username", "pulse@localhost")),
        timeout=config.smtp.timeout,
    )
    return replace(config, email_backend=settings.get("backend", "smtp"), smtp=smtp)


def display_email_settings(config: AppConfig) -> dict[str, Any]:
    settings = load_email_settings(config)
    if not settings:
        return {
            "provider": "gmail",
            "host": "smtp.gmail.com",
            "port": 587,
            "username": "",
            "from_address": "",
            "password_saved": False,
            "backend": config.email_backend,
        }
    return {
        **settings,
        "password": "",
        "password_saved": bool(settings.get("password")),
    }
