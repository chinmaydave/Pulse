from __future__ import annotations

from flask import Blueprint, current_app, flash, redirect, render_template, request, url_for

from .agents import ReminderAgent
from .email_service import email_service


bp = Blueprint("pulse", __name__)


def repository():
    return current_app.pulse_repository


def config():
    return current_app.config["PULSE_CONFIG"]


@bp.route("/")
def dashboard():
    records = repository().list_requests()
    return render_template(
        "dashboard.html",
        summary=repository().summary(),
        records=records,
        pending=repository().pending_for_reminder(config().reminder_days_ahead),
        workbook_path=config().workbook_path,
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
    mailer = email_service(cfg.use_outlook)
    agent = ReminderAgent(repository(), mailer, request.host_url.rstrip("/"))

    if request.method == "POST":
        request_id = request.form.get("request_id", "")
        if request_id == "all":
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
    messages = [(record, agent.build_message(record)) for record in pending]
    return render_template("reminders.html", messages=messages, days_ahead=cfg.reminder_days_ahead)


@bp.route("/reports")
def reports():
    return render_template(
        "reports.html",
        audit_rows=reversed(repository().audit_rows()),
        reminder_rows=reversed(repository().reminder_rows()),
    )
