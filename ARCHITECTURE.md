# Workflow

This product is an internal Python application for managing request records. It reads
from and writes back to an Excel workbook stored on SharePoint, surfaces forms and
dashboards to associates from a Python dev application, and automatically chases pending
requests over email. It is designed to run entirely on internal infrastructure with
no cloud dependencies.

The system is organized as a layered architecture. User-facing surfaces sit on top,
the processing core is in the middle, and the data repository is at the base.
Requests flow top‑down toward Excel; status, dashboards, and reports flow back up.

```mermaid
flowchart TD
    %% Start / high-level flow
    Start([Start: Project workflow])
    
    subgraph Backend["Backend development"]
      direction TB
      LocalExcel["Step 1a: Develop backend reading local Excel\n(rapid dev & tests)"]
      SharePointExcel["Step 1b: Switch backend to SharePoint Excel\n(production data source)"]
    end

    Core["Step 2: Build internal Python functions\n(data access, normalize rows, find upcoming expirations)"]

    subgraph ManagerUI["Manager workflows"]
      direction TB
      ManagerFront["Step 3: Manager frontend\n(view expirations, filter, review lists)"]
      TriggerRequests["Action: Send update requests to matched employees"]
    end

    subgraph EmployeeUI["Employee workflows"]
      direction TB
      EmployeeFront["Step 4: Employee frontend\n(update personal data, respond to requests)"]
      SubmitUpdate["Action: Submit changes -> triggers backend update"]
    end

    DataPersist["Step 5: Excel persistence & logging\n(update sheet; append notification logs, status, timestamps)"]

    subgraph Automation["Step 6: Agents (automation layer)"]
      direction LR
      Sync["Sync agents\n(sync Excel <-> app state)"]
      Reminder["Reminder agents\n(identify due items, send emails)"]
      Forms["Forms agents\n(apply form responses, status updates)"]
      Reporting["Reporting agents\n(audit & operational reports)"]
    end

    Comm["Communication layer\n(Outlook / email logging / deep links)"]
    Integrations["Step 7: Integrations\n(SharePoint, Microsoft Teams portal tab)"]
    End([Done / Deploy to internal infra])

    %% Primary linear flow
    Start --> LocalExcel --> SharePointExcel --> Core
    Core --> ManagerFront --> TriggerRequests --> EmployeeFront --> SubmitUpdate --> DataPersist
    DataPersist --> Automation
    Automation --> Comm --> DataPersist
    Automation --> Integrations --> End

    %% Cross links
    ManagerFront -- "optionally preview logs & agent suggestions" --> Reporting
    Reminder -- "sends notifications" --> Comm
    Sync -- "keeps data fresh" --> Core
    EmployeeFront -- "direct update triggers" --> Forms
    Forms -- "apply updates to Excel" --> DataPersist

    classDef step fill:#EEEDFE,stroke:#534AB7,color:#3C3489;
    classDef io fill:#E1F5EE,stroke:#0F6E56,color:#085041;
    classDef infra fill:#F1EFE8,stroke:#5F5E5A,color:#444441;
    class LocalExcel,SharePointExcel,Core,ManagerFront,EmployeeFront,DataPersist,Automation,Integrations step;
    class Comm,End io;
    class Sync,Reminder,Forms,Reporting infra;
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
