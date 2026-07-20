# Pulse

Internal Python development app for managing employee document expiration reminders from an Excel workbook.

The current MVP uses a OneDrive Excel sharing link as the active data source. It provides:

- Excel-backed expiration tracking with `Employees`, `AuditLog`, `ReminderLog`, and `ReminderRequests` sheets
- OneDrive Excel URL connection and refresh through the web interface
- Dashboard for expiration counts, overdue records, and reminder queue
- Manual and automatic expiration reminder preparation
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

The workbook is created at `data/EmployeeExpirations_OneDrive_Template.xlsx`. To use another workbook path:

```bash
PULSE_WORKBOOK_PATH=/path/to/requests.xlsx python -m pulse_app
```

## Automatic Reminder Agent

Pulse can run a background reminder agent inside the Flask process. The agent scans the Excel workbook, finds active requests due within the reminder window, sends reminders to associates, sends escalations to managers when a request is overdue with at least two prior reminders, and writes the reminder result back to Excel.

Automatic sending is opt-in so local test runs do not accidentally send real mail:

```bash
PULSE_EMAIL_BACKEND=outlook \
PULSE_AUTO_REMINDERS=true \
PULSE_APP_BASE_URL=http://SERVER_INTERNAL_IP:5000 \
python -m pulse_app
```

PowerShell:

```powershell
$env:PULSE_EMAIL_BACKEND="outlook"
$env:PULSE_AUTO_REMINDERS="true"
$env:PULSE_APP_BASE_URL="http://SERVER_INTERNAL_IP:5000"
python -m pulse_app
```

Useful settings:

```text
PULSE_REMINDER_DAYS_AHEAD=3
PULSE_REMINDER_SCAN_INTERVAL_SECONDS=300
PULSE_REMINDER_COOLDOWN_HOURS=24
PULSE_HOST=0.0.0.0
PULSE_PORT=5000
PULSE_DEBUG=false
```

Keep the active workbook closed in Excel while the agent is running. Excel creates a `~$...xlsx` lock file when the workbook is open, and that prevents Pulse from saving reminder counts and logs.

## Email Sending

Pulse defaults to Outlook sending. Users do not enter passwords or command-line email settings in Pulse.

```text
Email -> Send test email
Reminders -> Send
```

The receiver comes from the `Email` column. The sender/delegate address comes from the `Manager Email` column. The machine running Pulse must already be signed into Outlook. If Outlook does not allow the signed-in user to send on behalf of the manager mailbox, Outlook will reject the send.

Connect the workbook from the web UI:

```text
Data Source -> Use OneDrive workbook
```

Workbooks must be `.xlsx` files with an `Employees` sheet and these five columns:

```text
Title
Name
Email
Manager Email
Expirary Date
```

Pulse adds and maintains reminder/audit sheets in its active workbook copy. Use a OneDrive Excel sharing link that the app can access. Pulse tries normal share links, OneDrive shared-content download URLs, and embedded workbook download URLs. Fully private or sign-in-only OneDrive access should be handled later with Microsoft Graph authentication.

## Outlook Sending

By default reminders are logged to Excel instead of sent. On a Windows internal machine with Outlook installed, set either `PULSE_EMAIL_BACKEND=outlook` or the older compatibility flag:

```bash
PULSE_USE_OUTLOOK=true python -m pulse_app
```

The Outlook integration uses `pywin32`, which is only installed on Windows from `requirements.txt`.

## MVP Architecture

- `pulse_app/excel_repository.py`: reads employee expiration rows and writes reminder/audit rows
- `pulse_app/routes.py`: Flask web routes and form handling
- `pulse_app/agents.py`: first manual reminder agent implementation
- `pulse_app/email_service.py`: development logger and Outlook sender abstraction
- `pulse_app/onedrive_source.py`: downloads an accessible OneDrive Excel URL into the app
- `scripts/create_mock_data.py`: rebuilds the local mock Excel workbook

Future SharePoint, Teams, and autonomous agent work can plug into these boundaries without changing the user-facing workflow.
