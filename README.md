# google-sheets-recruitment-pipeline

# Serverless Recruitment & Workflow Automation Engine

An event-driven Google Apps Script (JavaScript) lifecycle automation system engineered on Google Workspace APIs to streamline multi-division applicant tracking, dynamic role resolution, and pipeline communications.

## Operational Overview

Scaling talent recruitment across multiple business divisions introduces significant operational friction:
* Manual status copying across spreadsheets causes communication delays and data drift.
* Off-the-shelf applicant tracking software (ATS) costs scale with seat counts and require continuous maintenance.
* Entry-level third-party webhooks (e.g., Zapier) face rate throttles, polling lag, and limited dynamic conditional routing.

This engine executes natively inside Google Cloud's Apps Script V8 runtime, reacting to spreadsheet edits instantly with zero server infrastructure overhead and zero recurring software licenses.

## Core Engineering Features

* **Dynamic Header Indexing:** Uses `indexOf()` to identify column positions (`Interview`, `Outcome`, `Email`, `Full Name`) at runtime, preventing script failures when team members add, delete, or rearrange columns.
* **Dynamic Division Mapping:** Automatically maps the active sheet tab name to standardized role titles across 10+ operational and research divisions without requiring separate trigger configurations per sheet.
* **Multi-Stage Event Handlers:** Separates early-stage interview dispatch (`careers@` sender) from formal onboarding packets and Drive folder authorizations (`contact@` sender).
* **Defensive Guard Clauses:** Excludes header row edits, blank email addresses, and unmonitored columns to maintain idempotent execution and prevent redundant email transmissions.

## Deployment Instructions

1. Open your Google Sheet tracking candidate records.
2. Navigate to **Extensions > Apps Script**.
3. Replace default contents with `Code.gs`.
4. Update the `PIPELINE_CONFIG` block at the top of `Code.gs` with your organization name, Calendly booking link, and target Google Drive resource folders.
5. In the Apps Script dashboard, navigate to **Triggers (⏰) > Add Trigger**:
   * **Function to run:** `handleHiringPipelineEdit`
   * **Deployment:** `Head`
   * **Event source:** `From spreadsheet`
   * **Event type:** `On edit`
6. Click **Save** and grant the required `SpreadsheetApp` and `GmailApp` OAuth scopes.

## Tech Stack
* **Language:** JavaScript (Google Apps Script V8 Engine)
* **APIs:** Google Sheets API, Gmail API, Google Drive API
* **Architecture:** Serverless Event-Driven Automations
