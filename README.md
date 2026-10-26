# AI Operations Task Automation

## Overview
A task management and automation workflow built using Google Sheets and Google Apps Script. The system helps track recurring operational tasks, record completion proof, and notify managers about overdue work.

## Features
- Centralized task master for recurring tasks
- Automated task-cycle creation based on task frequency
- Task completion tracking with proof or reference IDs
- Completion history maintained in a log
- Manager dashboard for task status and follow-up
- Email alerts for overdue tasks
- Escalation to leadership when manager acknowledgment is missing for 24 hours
- Automatic creation of the next task cycle after completion
- Time-driven automation using Apps Script triggers

## Technology Stack
- Google Sheets
- Google Apps Script
- JavaScript
- Gmail notifications

## Project Structure
- `Code.gs` — Automation logic
- `README.md` — Project documentation
- `screenshots/` — Screenshots of the workflow and dashboards

## Workflow
1. Tasks and their frequency are configured in the Task Master.
2. The system creates task cycles with due dates and deadlines.
3. Users mark tasks as completed and submit proof.
4. Completion records are stored in the Completion Log.
5. Overdue tasks trigger manager email alerts.
6. Unacknowledged overdue tasks can be escalated to leadership.
7. Completing a task generates the next scheduled cycle.

## Testing
The workflow was tested for task completion, proof recording, dashboard updates, email alerts, and next-cycle creation. Leadership escalation after 24 hours has been implemented but has not yet been fully verified end to end.

Some spreadsheet entries are test records and should not be interpreted as real operational activity.

## Setup
1. Create a Google Sheet with the required task and tracking tabs.
2. Open Extensions → Apps Script.
3. Add the code from `Code.gs`.
4. Configure task names, frequencies, deadlines, and authorized email addresses.
5. Run the setup function and authorize the required permissions.
6. Install the time-driven automation trigger.
7. Test the workflow using non-production data before operational use.

## Author
Rahul Dey

GitHub Portfolio: https://github.com/Rahul124679998
