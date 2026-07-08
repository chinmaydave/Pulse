from __future__ import annotations

from .config import AppConfig


def create_app(config: AppConfig | None = None):
    from flask import Flask

    from .excel_repository import ExcelRepository
    from .routes import bp

    app = Flask(__name__)
    app.config.from_mapping(
        SECRET_KEY="dev-pulse-secret",
        PULSE_CONFIG=config or AppConfig.from_env(),
    )

    app.pulse_repository = ExcelRepository(app.config["PULSE_CONFIG"].workbook_path)
    app.register_blueprint(bp)
    return app
