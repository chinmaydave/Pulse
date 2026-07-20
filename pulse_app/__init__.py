from __future__ import annotations

from .config import AppConfig


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
