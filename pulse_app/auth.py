from __future__ import annotations

import secrets

from flask import Blueprint, current_app, flash, redirect, render_template, request, session, url_for


bp = Blueprint("auth", __name__)


def is_authenticated() -> bool:
    return bool(session.get("authenticated"))


@bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        cfg = current_app.config["PULSE_CONFIG"]
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        username_ok = secrets.compare_digest(username, cfg.auth_username)
        password_ok = secrets.compare_digest(password, cfg.auth_password)
        if username_ok and password_ok:
            session.clear()
            session["authenticated"] = True
            return redirect(request.args.get("next") or url_for("pulse.dashboard"))
        flash("Invalid username or password.", "error")

    return render_template("login.html")


@bp.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))
