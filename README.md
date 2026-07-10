# Pulse

Internal Python development app for managing associate information update requests from an Excel workbook.

The current MVP intentionally uses a local mock workbook instead of SharePoint. It provides:

- Excel-backed request storage with `Requests`, `AuditLog`, and `ReminderLog` sheets
- Workbook upload through the web interface
- Dashboard for request counts, overdue work, and reminder queue
- Associate request forms that write submitted values back to Excel
- Manager status updates
- Manual reminder and escalation preparation
- Development email logging, with an optional Outlook sender path for Windows machines

## Run as an Internal Dev App

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/create_mock_data.py
python -m pulse_app
```

The app binds to `0.0.0.0:5000` by default so other users on the same internal network can open it with the server machine's IP address:

```text
http://SERVER_INTERNAL_IP:5000
```

On the server itself, this also works:

```text
http://127.0.0.1:5000
```

The workbook is created at `data/pulse_requests_mock.xlsx`. To use another workbook path:

```bash
PULSE_WORKBOOK_PATH=/path/to/requests.xlsx python -m pulse_app
```

You can also upload a workbook from the web UI:

```text
Data Source -> Upload workbook
```

Uploaded workbooks must be `.xlsx` files with a `Requests` sheet and the required request columns. If `AuditLog` or `ReminderLog` sheets are missing, the app creates them.

## Outlook Sending

By default reminders are logged to Excel instead of sent. On a Windows internal machine with Outlook installed, set:

```bash
PULSE_USE_OUTLOOK=true python -m pulse_app
```

The Outlook integration uses `pywin32`, which is only installed on Windows from `requirements.txt`.

## MVP Architecture

- `pulse_app/excel_repository.py`: reads and writes workbook rows
- `pulse_app/routes.py`: Flask web routes and form handling
- `pulse_app/agents.py`: first manual reminder agent implementation
- `pulse_app/email_service.py`: development logger and Outlook sender abstraction
- `scripts/create_mock_data.py`: rebuilds the local mock Excel workbook

Future SharePoint, Teams, and autonomous agent work can plug into these boundaries without changing the user-facing workflow.
