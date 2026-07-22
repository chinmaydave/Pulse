from __future__ import annotations

import json
import os
from pathlib import Path

from flask import Blueprint, current_app, flash, redirect, render_template, request, session, url_for

from .agents import ReminderAgent
from .email_service import email_service
from .excel_repository import ExcelRepository, EMPLOYEE_HEADERS
from .graph_onedrive import GraphAuthPending, complete_device_flow, download_private_workbook, start_device_flow
from .onedrive_source import download_onedrive_workbook
from .sender_credentials import (
    clear_gmail_app_password,
    gmail_sender_statuses,
    save_gmail_app_password,
)


bp = Blueprint("pulse", __name__)
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "pulse-admin"


@bp.before_app_request
def require_login():
    if request.endpoint in {"static", "pulse.login"}:
        return None
    if session.get("authenticated"):
        return None
    return redirect(url_for("pulse.login", next=request.full_path if request.query_string else request.path))


@bp.route("/login", methods=["GET", "POST"])
def login():
    if session.get("authenticated"):
        return redirect(request.args.get("next") or url_for("pulse.dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["authenticated"] = True
            session["username"] = username
            flash("Signed in.", "success")
            return redirect(request.args.get("next") or url_for("pulse.dashboard"))
        flash("Invalid username or password.", "error")

    return render_template("login.html")


@bp.post("/logout")
def logout():
    session.clear()
    flash("Signed out.", "success")
    return redirect(url_for("pulse.login"))


def repository():
    return current_app.pulse_repository


def config():
    return current_app.config["PULSE_CONFIG"]


def set_repository(workbook_path: Path) -> None:
    current_app.pulse_repository = ExcelRepository(workbook_path)
    if hasattr(current_app, "pulse_automatic_agent"):
        current_app.pulse_automatic_agent.repository = current_app.pulse_repository


def onedrive_source_path() -> Path:
    return config().upload_dir / "active_onedrive_url.txt"


def onedrive_source_meta_path() -> Path:
    return config().upload_dir / "active_onedrive_source.json"


def graph_token_path() -> Path:
    return config().upload_dir / "graph_token.json"


def graph_device_path() -> Path:
    return config().upload_dir / "graph_device_flow.json"


def read_json_file(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def write_private_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def read_onedrive_source() -> dict:
    source = read_json_file(onedrive_source_meta_path())
    if source.get("url"):
        return source
    source_url = read_onedrive_source_url()
    if source_url:
        return {"kind": "public", "url": source_url}
    return {"kind": "", "url": ""}


def read_onedrive_source_url() -> str:
    source_path = onedrive_source_path()
    if not source_path.exists():
        return ""
    return source_path.read_text(encoding="utf-8").strip()


def write_onedrive_source(kind: str, source_url: str) -> None:
    config().upload_dir.mkdir(parents=True, exist_ok=True)
    onedrive_source_path().write_text(source_url.strip(), encoding="utf-8")
    write_private_json(onedrive_source_meta_path(), {"kind": kind, "url": source_url.strip()})


def graph_status() -> dict:
    token = read_json_file(graph_token_path())
    device = read_json_file(graph_device_path())
    return {
        "connected": bool(token.get("refresh_token") or token.get("access_token")),
        "client_id": token.get("client_id") or device.get("client_id", ""),
        "tenant": token.get("tenant") or device.get("tenant", "common"),
        "pending": bool(device.get("device_code")),
        "user_code": device.get("user_code", ""),
        "verification_uri": device.get("verification_uri", ""),
        "message": device.get("message", ""),
    }


def activate_onedrive_workbook(source_url: str, kind: str = "public") -> Path:
    cfg = config()
    cfg.upload_dir.mkdir(parents=True, exist_ok=True)
    uploaded_path = cfg.upload_dir / "onedrive_expirations.xlsx"
    pending_path = cfg.upload_dir / "onedrive_expirations.pending.xlsx"
    pending_path.unlink(missing_ok=True)
    if kind == "private":
        token_cache = read_json_file(graph_token_path())
        updated_cache = download_private_workbook(source_url, pending_path, token_cache)
        write_private_json(graph_token_path(), updated_cache)
    else:
        download_onedrive_workbook(source_url, pending_path)
    ExcelRepository.validate_workbook(pending_path)
    pending_path.replace(uploaded_path)
    write_onedrive_source(kind, source_url)
    set_repository(uploaded_path)
    return uploaded_path


def automatic_agent():
    return current_app.pulse_automatic_agent


@bp.route("/")
def dashboard():
    records = repository().list_employees()
    return render_template(
        "dashboard.html",
        summary=repository().summary(),
        records=records,
        pending=repository().pending_for_reminder(config().reminder_days_ahead),
        workbook_path=repository().workbook_path,
        onedrive_source_url=read_onedrive_source_url(),
        agent_snapshot=automatic_agent().snapshot(),
    )


@bp.route("/data-source", methods=["GET", "POST"])
def data_source():
    if request.method == "POST":
        action = request.form.get("action", "connect")
        if action == "refresh-onedrive":
            source = read_onedrive_source()
            source_url = source.get("url", "")
            if not source_url:
                flash("Paste and connect a OneDrive Excel URL before refreshing.", "error")
                return redirect(url_for("pulse.data_source"))
            try:
                activate_onedrive_workbook(source_url, source.get("kind", "public"))
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            flash("OneDrive workbook refreshed from the saved link.", "success")
            return redirect(url_for("pulse.dashboard"))

        if action == "start-private-onedrive":
            client_id = request.form.get("graph_client_id", "").strip()
            tenant = request.form.get("graph_tenant", "common").strip() or "common"
            source_url = request.form.get("private_onedrive_url", "").strip()
            if not source_url:
                flash("Paste a private OneDrive Excel URL.", "error")
                return redirect(url_for("pulse.data_source"))
            try:
                device = start_device_flow(client_id, tenant)
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            device["client_id"] = client_id
            device["tenant"] = tenant
            device["source_url"] = source_url
            write_private_json(graph_device_path(), device)
            flash("Microsoft sign-in started. Use the code shown below, then click Finish private connection.", "success")
            return redirect(url_for("pulse.data_source"))

        if action == "finish-private-onedrive":
            device = read_json_file(graph_device_path())
            if not device.get("device_code"):
                flash("Start Microsoft sign-in before finishing the private connection.", "error")
                return redirect(url_for("pulse.data_source"))
            try:
                token = complete_device_flow(device["client_id"], device["tenant"], device["device_code"])
                token["client_id"] = device["client_id"]
                token["tenant"] = device["tenant"]
                write_private_json(graph_token_path(), token)
                graph_device_path().unlink(missing_ok=True)
                activate_onedrive_workbook(device["source_url"], "private")
            except GraphAuthPending as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            flash("Private OneDrive workbook connected and activated.", "success")
            return redirect(url_for("pulse.dashboard"))

        source_url = request.form.get("public_onedrive_url", "").strip()
        if source_url:
            try:
                activate_onedrive_workbook(source_url, "public")
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            flash("Public OneDrive workbook connected and activated.", "success")
            return redirect(url_for("pulse.dashboard"))

        flash("Paste a OneDrive Excel URL.", "error")
        return redirect(url_for("pulse.data_source"))

    source = read_onedrive_source()
    return render_template(
        "data_source.html",
        workbook_path=repository().workbook_path,
        onedrive_source_url=source.get("url", ""),
        onedrive_source_kind=source.get("kind", ""),
        graph_status=graph_status(),
        required_headers=", ".join(EMPLOYEE_HEADERS),
    )


@bp.route("/reminders", methods=["GET", "POST"])
def reminders():
    cfg = config()
    mailer = email_service(cfg)
    agent = ReminderAgent(repository(), mailer, request.host_url.rstrip("/"))
    expiration_filter = request.values.get("expiration_filter", "all")

    if request.method == "POST":
        action = request.form.get("action", "")
        if action == "save-gmail-sender":
            try:
                save_gmail_app_password(
                    cfg,
                    request.form.get("gmail_sender", ""),
                    request.form.get("gmail_app_password", ""),
                )
            except ValueError as exc:
                flash(str(exc), "error")
            else:
                flash("Gmail sender password saved for reminder sending.", "success")
            return redirect(url_for("pulse.reminders", expiration_filter=expiration_filter))

        if action == "clear-gmail-sender":
            sender = request.form.get("gmail_sender", "")
            clear_gmail_app_password(cfg, sender)
            flash(f"Gmail sender password cleared for {sender}.", "success")
            return redirect(url_for("pulse.reminders", expiration_filter=expiration_filter))

        target_key = request.form.get("target_key", "")
        if target_key == "automatic-once":
            results = automatic_agent().run_once()
            flash(f"Automatic agent processed {len(results)} reminder(s).", "success")
        elif target_key == "all":
            results = agent.send_pending(
                cfg.reminder_days_ahead,
                None if expiration_filter == "all" else expiration_filter,
            )
            sent_count = sum(1 for result in results if result["status"] == "sent")
            skipped_count = sum(1 for result in results if result["status"] == "skipped")
            flash(f"Processed {len(results)} pending reminders: {sent_count} sent, {skipped_count} skipped.", "success")
        else:
            target = repository().find_target_by_key(target_key)
            if target is None:
                flash(f"Reminder target {target_key} was not found.", "error")
            else:
                result = agent.send_for_target(target)
                category = "success" if result["status"] in {"sent", "skipped"} else "error"
                flash(
                    f"Reminder for {target.employee_name} ({target.field_name}): "
                    f"{result['status']} to {result['recipient']}. {result['detail']}",
                    category,
                )
        return redirect(url_for("pulse.reminders", expiration_filter=expiration_filter))

    pending = repository().pending_for_reminder(
        cfg.reminder_days_ahead,
        None if expiration_filter == "all" else expiration_filter,
    )
    automatic_due = agent.due_for_automatic_send(
        cfg.reminder_days_ahead,
        cfg.reminder_cooldown_hours,
        pending,
    )
    filter_label = {
        "all": "All expiring records",
        "overdue": "Overdue",
        "30": "Expires in 30 days or less",
        "14": "Expires in 14 days or less",
        "7": "Expires in 7 days or less",
    }.get(expiration_filter, "All expiring records")
    messages = [(target, agent.build_message(target), repository().last_sent_for_target(target)) for target in pending]
    return render_template(
        "reminders.html",
        messages=messages,
        gmail_sender_statuses=gmail_sender_statuses(cfg, pending),
        automatic_due=automatic_due,
        days_ahead=cfg.reminder_days_ahead,
        cooldown_hours=cfg.reminder_cooldown_hours,
        agent_snapshot=automatic_agent().snapshot(),
        expiration_filter=expiration_filter,
        filter_label=filter_label,
    )


@bp.route("/reports")
def reports():
    return render_template(
        "reports.html",
        audit_rows=reversed(repository().audit_rows()),
        reminder_rows=reversed(repository().reminder_rows()),
    )
