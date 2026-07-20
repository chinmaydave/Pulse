from __future__ import annotations

import atexit
import os

from .agents import AutomaticReminderAgent
from .config import AppConfig
from .email_service import email_service


def create_app(config: AppConfig | None = None):
    from flask import Flask

    from .routes import bp

    app = Flask(__name__)
    cfg = config or AppConfig.from_env()
    app.config.from_mapping(
        SECRET_KEY="dev-pulse-secret",
        PULSE_CONFIG=cfg,
    )

    app.pulse_repository = _build_repository(cfg)
    app.pulse_automatic_agent = _build_automatic_agent(app)
    if _should_start_background_agents(app.config["PULSE_CONFIG"]):
        app.pulse_automatic_agent.start()
        atexit.register(app.pulse_automatic_agent.stop)

    app.register_blueprint(bp)
    return app


def _build_repository(cfg: AppConfig):
    """Return the correct repository backend based on cfg.data_backend."""
    if cfg.data_backend == "onedrive":
        from pathlib import Path
        from .onedrive_repository import OneDriveRepository

        def _stub_download(dest_path: Path) -> None:
            raise NotImplementedError(
                "OneDrive auth not yet implemented — complete Part 1a "
                "and supply real download_fn / upload_fn to OneDriveRepository."
            )

        def _stub_upload(src_path: Path) -> None:
            raise NotImplementedError(
                "OneDrive auth not yet implemented — complete Part 1a "
                "and supply real download_fn / upload_fn to OneDriveRepository."
            )

        return OneDriveRepository(
            cache_path=cfg.workbook_path,
            download_fn=_stub_download,
            upload_fn=_stub_upload,
        )

    from .excel_repository import ExcelRepository
    return ExcelRepository(cfg.workbook_path)


def _build_automatic_agent(app):
    cfg = app.config["PULSE_CONFIG"]
    return AutomaticReminderAgent(
        repository=app.pulse_repository,
        mailer=email_service(cfg),
        app_base_url=cfg.app_base_url,
        email_backend=cfg.email_backend,
        use_outlook=cfg.use_outlook,
        days_ahead=cfg.reminder_days_ahead,
        scan_interval_seconds=cfg.reminder_scan_interval_seconds,
        cooldown_hours=cfg.reminder_cooldown_hours,
        enabled=cfg.auto_reminders_enabled,
    )


def _should_start_background_agents(cfg: AppConfig) -> bool:
    if not cfg.auto_reminders_enabled:
        return False
    if not cfg.debug:
        return True
    return os.environ.get("WERKZEUG_RUN_MAIN") == "true"
