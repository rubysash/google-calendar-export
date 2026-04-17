# Google Calendar Data Extractor

A Python-based utility to extract Google Calendar events and export them to a formatted Excel spreadsheet.

![Demo](https://github.com/rubysash/google-calendar-export/blob/main/demo.png?raw=true)

## Project Overview

*   **Technology Stack:** Python 3.12+, Google Calendar API v3, `pandas`, `openpyxl`.
*   **Key Features:**
    *   Extracts event details, attendees, meeting links, and metadata.
    *   **Data Preservation:** Keeps original formatting and newlines in event descriptions.
    *   **Smart Extraction:** Uses regex to pull email addresses and phone numbers into dedicated columns while keeping them in the description.
    *   **Smart Deduplication:** Deduplicates phone numbers based on numeric content.
    *   Outputs a styled Excel file with auto-adjusted column widths and filters (uses comma separators for multiple values).

---

## 1. Google Cloud Console Setup (OAuth 2.0)

This project requires a `credentials.json` file from the Google Cloud Console to access the Calendar API.

1.  **Create Project:** Go to the [Google Cloud Console](https://console.cloud.google.com/), click the project dropdown (top left) > **New Project**, name it, and click **Create**.
2.  **Enable API:** Navigate to **APIs & Services > Library**, search for "Google Calendar API", and click **Enable**.
3.  **OAuth Consent Screen:**
    *   Go to **APIs & Services > OAuth consent screen**.
    *   Select **User Type: External** and click **Create**.
    *   **App Information:** Fill in "App name", "User support email", and "Developer contact info". Click **Save and Continue**.
    *   **Scopes:** Click **Add or Remove Scopes**. Search for `.../auth/calendar.readonly`, check the box, and click **Update**. Click **Save and Continue**.
    *   **Test Users:** (CRITICAL) Click **+ ADD USERS**. Enter the email address of the account whose calendar you want to extract.
4.  **Create Credentials:**
    *   Go to **APIs & Services > Credentials**.
    *   Click **+ CREATE CREDENTIALS** > **OAuth client ID**. (Do NOT select "API key").
    *   Select **Application type: Desktop app**.
    *   Name it and click **Create**.
    *   In the dialog, click **DOWNLOAD JSON**.
    *   Rename this file to `credentials.json` and move it into the project folder.

---

## 2. Installation & Environment Setup

Follow these steps to set up the project locally:

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/rubysash/google-calendar-export.git
    cd google-calendar-export
    ```

2.  **Create and activate the virtual environment:**
    ```bash
    # Create the venv with the specific project name
    python -m venv google-calendar-export

    # Activate on Windows:
    google-calendar-export\Scripts\activate

    # Activate on macOS/Linux:
    source google-calendar-export/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    python -m pip install pip --upgrade pip
    pip install -r requirements.txt
    ```

---

## 3. Usage

### Basic Command
Extract events from the last 45 days:
```bash
python main.py --days 45
```

### Options & Flags
*   `--days NUMBER`: **(Required)** Number of days in the past to look back.
    *   *Limit:* There is no hard limit on how far back you can go, but very large ranges (e.g., thousands of days) may take longer to process and hit API pagination limits.
*   `--output FILENAME`: **(Optional)** Custom name for the Excel file.
    *   *Default:* `YYYY-MM-DD-Nd-calendar_export.xlsx` (e.g., `2026-04-17-45d-calendar_export.xlsx`).
*   `--reauth`: **(Optional)** Deletes the local `token.pickle` and forces a new browser login. Use this to switch to a different Google account.

### Examples

**Quarterly Report (90 days):**
```bash
python main.py --days 90 --output quarterly_report.xlsx
```

**Switch Accounts & Quick Export (7 days):**
```bash
python main.py --reauth --days 7
```

**Full Year Audit:**
```bash
python main.py --days 365 --output 2025_full_year.xlsx
```

---

## Important Notes

*   **Authentication:** After the first run, a `token.pickle` file is created. This stores your login session.
*   **Whose data?** The script pulls data from the account that completes the OAuth login in the browser.
*   **Test Users:** While the app is in "Testing" mode in Google Cloud Console, only emails added to the "Test users" list can authenticate.
