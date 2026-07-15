from __future__ import annotations

from pathlib import Path
from shutil import copyfileobj

from flask import Blueprint, current_app, flash, redirect, render_template, request, url_for
from werkzeug.utils import secure_filename

from .agents import ReminderAgent
from .email_service import email_service
from .excel_repository import ExcelRepository, REQUEST_HEADERS


bp = Blueprint("pulse", __name__)


def repository():
    return current_app.pulse_repository


def config():
    return current_app.config["PULSE_CONFIG"]


def set_repository(workbook_path: Path) -> None:
    current_app.pulse_repository = ExcelRepository(workbook_path)
    if hasattr(current_app, "pulse_automatic_agent"):
        current_app.pulse_automatic_agent.repository = current_app.pulse_repository


def automatic_agent():
    return current_app.pulse_automatic_agent


@bp.route("/")
def dashboard():
    records = repository().list_requests()
    return render_template(
        "dashboard.html",
        summary=repository().summary(),
        records=records,
        pending=repository().pending_for_reminder(config().reminder_days_ahead),
        workbook_path=repository().workbook_path,
        agent_snapshot=automatic_agent().snapshot(),
    )


@bp.route("/data-source", methods=["GET", "POST"])
def data_source():
    cfg = config()
    if request.method == "POST":
        upload = request.files.get("workbook")
        if not upload or not upload.filename:
            flash("Choose an Excel workbook to upload.", "error")
            return redirect(url_for("pulse.data_source"))
        if not upload.filename.lower().endswith(".xlsx"):
            flash("Upload a .xlsx workbook.", "error")
            return redirect(url_for("pulse.data_source"))

        cfg.upload_dir.mkdir(parents=True, exist_ok=True)
        filename = secure_filename(upload.filename)
        uploaded_path = cfg.upload_dir / filename
        with uploaded_path.open("wb") as target:
            copyfileobj(upload.stream, target)

        try:
            ExcelRepository.validate_workbook(uploaded_path)
            set_repository(uploaded_path)
        except Exception as exc:
            uploaded_path.unlink(missing_ok=True)
            flash(str(exc), "error")
            return redirect(url_for("pulse.data_source"))

        flash(f"Workbook uploaded and activated: {filename}", "success")
        return redirect(url_for("pulse.dashboard"))

    return render_template(
        "data_source.html",
        workbook_path=repository().workbook_path,
        required_headers=", ".join(REQUEST_HEADERS),
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

    if request.method == "POST":
        request_id = request.form.get("request_id", "")
        if request_id == "automatic-once":
            results = automatic_agent().run_once()
            flash(f"Automatic agent processed {len(results)} reminder(s).", "success")
        elif request_id == "all":
            results = agent.send_pending(cfg.reminder_days_ahead)
            flash(f"Processed {len(results)} pending reminders.", "success")
        else:
            record = repository().get_request(request_id)
            if record is None:
                flash(f"Request {request_id} was not found.", "error")
            else:
                result = agent.send_for_record(record)
                flash(f"Reminder for {request_id}: {result['status']} ({result['recipient']}).", "success")
        return redirect(url_for("pulse.reminders"))

    pending = repository().pending_for_reminder(cfg.reminder_days_ahead)
    automatic_due = agent.due_for_automatic_send(cfg.reminder_days_ahead, cfg.reminder_cooldown_hours)
    messages = [(record, agent.build_message(record)) for record in pending]
    return render_template(
        "reminders.html",
        messages=messages,
        automatic_due=automatic_due,
        days_ahead=cfg.reminder_days_ahead,
        cooldown_hours=cfg.reminder_cooldown_hours,
        agent_snapshot=automatic_agent().snapshot(),
    )


@bp.route("/reports")
def reports():
    return render_template(
        "reports.html",
        audit_rows=reversed(repository().audit_rows()),
        reminder_rows=reversed(repository().reminder_rows()),
    )
