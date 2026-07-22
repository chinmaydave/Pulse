# Pulse

Internal Python development app for managing employee document expiration reminders from an Excel workbook.

The current MVP uses a OneDrive Excel sharing link as the active data source. It provides:

- Excel-backed expiration tracking with `Employees`, `AuditLog`, `ReminderLog`, and `ReminderRequests` sheets
- Public OneDrive link download and private OneDrive Microsoft Graph sign-in through the web interface
- Dashboard for expiration counts, overdue records, and reminder queue
- Manual and automatic expiration reminder preparation
- Outlook reminder sending from Excel email addresses

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

Pulse runs a background reminder agent inside the Flask process. The agent scans the Excel workbook, finds records due within the reminder window, sends Outlook reminders, and writes reminder results back to Excel.

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

Pulse chooses the sender based on the Excel `Manager Email` column. There is no Email Settings tab and no command-line email setup.

```text
Reminders -> Send
```

The receiver comes from the `Email` column. The sender comes from the `Manager Email` column.

- Gmail manager senders use Gmail SMTP on Mac and Windows. Pulse asks for that Gmail account's app password on the Reminders page the first time it sees the sender.
- Non-Gmail manager senders use Outlook desktop. The machine running Pulse must already be signed into Outlook, and Outlook must allow send-on-behalf for that manager mailbox.

Connect the workbook from the web UI.

For public links:

```text
Data Source -> Public OneDrive link -> Use public workbook
```

For private links:

```text
Data Source -> Private OneDrive link
Paste the private Excel URL
Enter a Microsoft app client ID
Use tenant common, consumers, organizations, or your tenant ID
Click Start Microsoft sign-in
Open the Microsoft device login link and enter the code
Return to Pulse and click Finish private connection
```

Workbooks must be `.xlsx` files with an `Employees` sheet and these five columns:

```text
Title
Name
Email
Manager Email
Expirary Date
```

Pulse adds and maintains reminder/audit sheets in its active workbook copy. Public links use direct download attempts. Private links use Microsoft Graph device-code sign-in and store the delegated token locally on the Pulse machine under `data/uploads/graph_token.json`.

## MVP Architecture

- `pulse_app/excel_repository.py`: reads employee expiration rows and writes reminder/audit rows
- `pulse_app/routes.py`: Flask web routes and form handling
- `pulse_app/agents.py`: first manual reminder agent implementation
- `pulse_app/email_service.py`: Excel-routed Gmail/Outlook sender abstraction
- `pulse_app/onedrive_source.py`: downloads an accessible OneDrive Excel URL into the app
- `pulse_app/graph_onedrive.py`: signs into Microsoft Graph and downloads private OneDrive workbooks
- `scripts/create_mock_data.py`: rebuilds the local mock Excel workbook

Future SharePoint, Teams, and autonomous agent work can plug into these boundaries without changing the user-facing workflow.
