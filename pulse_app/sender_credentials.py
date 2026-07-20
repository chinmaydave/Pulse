from __future__ import annotations

import json
from pathlib import Path

from .config import AppConfig
from .models import ExpirationTarget


def is_gmail_address(email: str) -> bool:
    domain = email.rsplit("@", 1)[-1].strip().lower()
    return domain in {"gmail.com", "googlemail.com"}


def credential_path(config: AppConfig) -> Path:
    return config.upload_dir / "sender_credentials.json"


def load_sender_credentials(config: AppConfig) -> dict[str, str]:
    path = credential_path(config)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return {str(key).lower(): str(value) for key, value in data.items()}


def save_gmail_app_password(config: AppConfig, sender: str, app_password: str) -> None:
    sender = sender.strip().lower()
    app_password = app_password.strip().replace(" ", "")
    if not sender or not is_gmail_address(sender) or not app_password:
        raise ValueError("Enter a Gmail sender address and app password.")

    config.upload_dir.mkdir(parents=True, exist_ok=True)
    credentials = load_sender_credentials(config)
    credentials[sender] = app_password
    credential_path(config).write_text(json.dumps(credentials, indent=2), encoding="utf-8")


def gmail_senders_missing_credentials(config: AppConfig, targets: list[ExpirationTarget]) -> list[str]:
    credentials = load_sender_credentials(config)
    senders = {
        target.manager_email.strip().lower()
        for target in targets
        if target.manager_email and is_gmail_address(target.manager_email)
    }
    return sorted(sender for sender in senders if sender not in credentials)
