from __future__ import annotations

import atexit
import os

from .agents import AutomaticReminderAgent
from .config import AppConfig
from .email_service import email_service


def create_app(config: AppConfig | None = None):
    from flask import Flask, redirect, request, url_for

    from .auth import bp as auth_bp, is_authenticated
    from .excel_repository import ExcelRepository
    from .routes import bp

    app = Flask(__name__)
    app.config.from_mapping(
        SECRET_KEY="dev-pulse-secret",
        PULSE_CONFIG=config or AppConfig.from_env(),
    )

    app.pulse_repository = ExcelRepository(app.config["PULSE_CONFIG"].workbook_path)
    app.pulse_automatic_agent = _build_automatic_agent(app)
    if _should_start_background_agents(app.config["PULSE_CONFIG"]):
        app.pulse_automatic_agent.start()
        atexit.register(app.pulse_automatic_agent.stop)

    app.register_blueprint(auth_bp)
    app.register_blueprint(bp)

    @app.before_request
    def _require_login():
        if request.endpoint == "static" or (request.endpoint or "").startswith("auth."):
            return None
        if not is_authenticated():
            return redirect(url_for("auth.login", next=request.path))
        return None

    return app


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
