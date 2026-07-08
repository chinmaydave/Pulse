# Architecture

Task Portal is an internal Python application for managing request records. It reads
from and writes back to an Excel workbook stored on SharePoint, surfaces forms and
dashboards to associates inside Microsoft Teams, and automatically chases pending
requests over email. It is designed to run entirely on internal infrastructure with
no cloud dependencies.

The system is organized as a layered architecture. User-facing surfaces sit on top,
the processing core is in the middle, and the data repository is at the base.
Requests flow top‑down toward Excel; status, dashboards, and reports flow back up.

```mermaid
flowchart TD
    Teams["<b>Microsoft Teams</b><br/>SPFx tab app — embedded primary UI"]
    Portal["<b>User access layer</b><br/>Web portal — forms, status, dashboards"]
    App["<b>Application layer</b><br/>Python web app — business logic, request processing"]

    subgraph Automation["Automation layer — independent Python agents"]
        direction LR
        Sync["<b>Sync</b><br/>syncs Excel"]
        Reminder["<b>Reminder</b><br/>sends emails"]
        Forms["<b>Forms</b><br/>updates status"]
        Reporting["<b>Reporting</b><br/>audit reports"]
    end

    Comm["<b>Communication layer</b><br/>Outlook via PyWin32 — reminder and escalation emails"]
    Data["<b>Data layer</b><br/>SharePoint Excel — central request repository"]

    Teams --> Portal --> App --> Automation
    Automation --> Comm --> Data

    classDef teal fill:#E1F5EE,stroke:#0F6E56,color:#085041;
    classDef purple fill:#EEEDFE,stroke:#534AB7,color:#3C3489;
    classDef gray fill:#F1EFE8,stroke:#5F5E5A,color:#444441;

    class Teams,Portal teal;
    class App,Sync,Reminder,Forms,Reporting purple;
    class Comm,Data gray;
    style Automation fill:none,stroke:#73726c,stroke-dasharray:5 4,color:#73726c;
```

## Layers

**Data layer.** A SharePoint-hosted Excel workbook is the single source of truth for
request records, user details, due dates, statuses, and reminder tracking. End users
do not touch the file directly — only designated administrators or managers maintain
source records.

**Application layer.** A Python web application holds the business logic: request
processing, form submissions, and writing updates back to Excel.

**Automation layer.** Four independent Python agents each own one responsibility:

| Agent | Responsibility |
| --- | --- |
| Sync | Reads and synchronizes records from Excel. |
| Reminder | Identifies pending requests, composes messages, and sends automated reminder emails. |
| Forms | Updates request status and stores user responses. |
| Reporting | Generates operational and audit reports. |

Keeping these as separate agents is also where the longer-term "agentic AI" automation
is intended to plug in.

**Communication layer.** Integrates with Microsoft Outlook via Python
(PyWin32 / Outlook client libraries). The Reminder agent sends emails to associates
and escalation notifications when required; email links deep-link the user straight to
the relevant form in the portal.

**User access layer.** An internal web portal lets users view assigned requests,
complete forms, track submission status, and access dashboards and reports.

**Microsoft Teams integration.** The portal is published as a Microsoft Teams tab
application, built with the SharePoint Framework (SPFx) as a custom installable app.
Users reach the solution directly inside Teams with no separate URL. Teams is the
primary interface, while all processing stays within the Python application and
internal infrastructure, and SharePoint Excel remains the underlying data store.

## Team ownership

| Owner | Area |
| --- | --- |
| Colin Bertrand | Agentic AI development |
| Chinmay Dave | Python development |
| Liam Ben-Zvi | Outlook and Teams app integration |
| Kavit Timbadia | Testing, documentation |

## Editing this diagram

The diagram above is a [Mermaid](https://mermaid.js.org/) `flowchart`, rendered
natively by GitHub. To change it, edit the fenced ```mermaid``` block — add a node,
rename a label, or re-wire an arrow — no image tooling required. A rendered
`docs/architecture.svg` is also kept in the repo for contexts that don't render
Mermaid (e.g. some slide tools and PDF exports).
