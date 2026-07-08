# Pulse

### Team Members
Colin Bertrand - Agentic AI Development

Chinmay Dave - Python Development

Liam Ben-Zvi - Integration to Outlook and Teams App

Kavit Timbadia - UI, Testing, Documentation

---

### Project Scope

Build an internal Python application that:

- Reads requests from an Excel file stored on SharePoint
- Provides a web interface for associates to complete forms to update information
- Update form information back to Excel sheet
- Automatically sends reminder emails for pending requests
- Updates status back to Excel
- Runs completely on internal machine/server with no cloud dependencies. (start as a python dev application)
- Start with manual selection and sending, eventually automate processes with agentic AI

---

### Project Architecture

1. **Data Layer**

- SharePoint Excel serves as the central repository for requests, records, user details, due dates, statuses, and reminder tracking

2. **Application Layer**

- Standalone Python web application used to run and manage the request workflow
- Built for internal use and hosted on an internal machine/server
- Provides the user interface that associates can access
- Handles backend processing including:
  - Readin request data from SharePoint Excel
  - Writing completed form data back to Excel
  - Updating request statuses
  - Triggering reminder and escalation emails
- Uses Microsoft APIs and/or approved Outlook libraries, for Excel and email integration
- Designed as the MVP foundation before adding automation from AI Agents and additional integration like microsoft Teams

3. **Automation Layer**

      Independent AI/Python agents perform specific functions:

- `Sync Agents`: Reads and synchronizes records from Excel
- `Reminder Agents`: Identifies pending requests, creates messages, and sends automated reminder emails
- `Form Processing Agents`: Updates request status and stores user responses
- `Reporting Agents`: Generates operational and audit reports

    Agents will be configured after a working MVP is developed. Multiple agents may exist for each section stated above.

4. **Communication Layer**

- Integration with **Microsoft Product**
  - Microsoft Outlook using Python (PyWin32/Outlook client libraries)
- Reminder Agent automatically sends emails/messages to associates and escalation notifications when required
- Email/message links direct users to the appropriate form within the application

5. **User Access Layer**

- Internal web portal allowing users to:
  - View assigned reuests
  - Complete forms
  - Track submission status
  - Access dashboards and reports
 

End users will interact with the solution through a web-based interface. 
The Python application will provide the forms, dashboards, reporting, and workflow functionality, 
while SharePoint Excel will remain the underlying data repository. 
Users will not interact directly with the Excel file except for designated administrators or managers maintaining source records.
 

