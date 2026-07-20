from __future__ import annotations

from pathlib import Path

from flask import Blueprint, current_app, flash, redirect, render_template, request, url_for

from .agents import ReminderAgent
from .email_service import email_service
from .excel_repository import ExcelRepository, EMPLOYEE_HEADERS
from .onedrive_source import download_onedrive_workbook
from .sender_credentials import (
    clear_gmail_app_password,
    gmail_sender_statuses,
    save_gmail_app_password,
)


bp = Blueprint("pulse", __name__)


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


def read_onedrive_source_url() -> str:
    source_path = onedrive_source_path()
    if not source_path.exists():
        return ""
    return source_path.read_text(encoding="utf-8").strip()


def activate_onedrive_workbook(source_url: str) -> Path:
    cfg = config()
    cfg.upload_dir.mkdir(parents=True, exist_ok=True)
    uploaded_path = cfg.upload_dir / "onedrive_expirations.xlsx"
    pending_path = cfg.upload_dir / "onedrive_expirations.pending.xlsx"
    pending_path.unlink(missing_ok=True)
    download_onedrive_workbook(source_url, pending_path)
    ExcelRepository.validate_workbook(pending_path)
    pending_path.replace(uploaded_path)
    onedrive_source_path().write_text(source_url.strip(), encoding="utf-8")
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
            source_url = read_onedrive_source_url()
            if not source_url:
                flash("Paste and connect a OneDrive Excel URL before refreshing.", "error")
                return redirect(url_for("pulse.data_source"))
            try:
                activate_onedrive_workbook(source_url)
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            flash("OneDrive workbook refreshed from the saved link.", "success")
            return redirect(url_for("pulse.dashboard"))

        source_url = request.form.get("onedrive_url", "").strip()
        if source_url:
            try:
                activate_onedrive_workbook(source_url)
            except Exception as exc:
                flash(str(exc), "error")
                return redirect(url_for("pulse.data_source"))

            flash("OneDrive workbook connected and activated.", "success")
            return redirect(url_for("pulse.dashboard"))

        flash("Paste a OneDrive Excel URL.", "error")
        return redirect(url_for("pulse.data_source"))

    return render_template(
        "data_source.html",
        workbook_path=repository().workbook_path,
        onedrive_source_url=read_onedrive_source_url(),
        required_headers=", ".join(EMPLOYEE_HEADERS),
    )


@bp.route("/requests")
def requests_index():
    status = request.args.get("status", "")
    associate = request.args.get("associate", "")
    records = repository().list_requests()
    if status:
        records = [record for record in records if record.status == status]
    if associate:
        records = [record for record in records if associate.lower() in record.assigned_to.lower()]
    return render_template("requests.html", records=records, status=status, associate=associate)


@bp.route("/requests/<request_id>", methods=["GET", "POST"])
def request_detail(request_id: str):
    record = repository().get_request(request_id)
    if record is None:
        flash(f"Request {request_id} was not found.", "error")
        return redirect(url_for("pulse.requests_index"))

    if request.method == "POST":
        submitted_value = request.form.get("submitted_value", "").strip()
        notes = request.form.get("notes", "").strip()
        actor = request.form.get("actor", record.assigned_to).strip() or record.assigned_to
        if not submitted_value:
            flash("Submitted value is required.", "error")
            return render_template("request_detail.html", record=record)
        repository().submit_response(request_id, submitted_value, notes, actor)
        flash(f"{request_id} was submitted and written back to Excel.", "success")
        return redirect(url_for("pulse.request_detail", request_id=request_id))

    return render_template("request_detail.html", record=record)


@bp.post("/requests/<request_id>/status")
def update_status(request_id: str):
    status = request.form.get("status", "")
    actor = request.form.get("actor", "manager").strip() or "manager"
    repository().update_status(request_id, status, actor, f"Manual status update from web portal.")
    flash(f"{request_id} status updated to {status}.", "success")
    return redirect(url_for("pulse.request_detail", request_id=request_id))


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
            results = agent.send_pending(cfg.reminder_days_ahead)
            flash(f"Processed {len(results)} pending reminders.", "success")
        else:
            target = repository().find_target_by_key(target_key)
            if target is None:
                flash(f"Reminder target {target_key} was not found.", "error")
            else:
                result = agent.send_for_target(target)
                category = "success" if result["status"] == "sent" else "error"
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
    messages = [(target, agent.build_message(target)) for target in pending]
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
