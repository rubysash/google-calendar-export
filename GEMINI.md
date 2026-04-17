# Google Calendar Data Extractor

A Python-based utility to extract Google Calendar events and export them to a formatted Excel spreadsheet.

## Project Overview

*   **Technology Stack:** Python 3.12+, Google Calendar API v3, `pandas`, `openpyxl`.
*   **Core Functionality:** The script pulls data from the **authenticated user's primary calendar**. It is designed so that different users can run the same script (using the same `credentials.json`) to extract their own calendar data.
*   **Key Features:**
    *   Extracts event details, attendees, meeting links, and metadata.
    *   Uses regex to pull email addresses and phone numbers from event descriptions/locations.
    *   Outputs a styled Excel file with auto-adjusted column widths and filters.

## Setup and Requirements

1.  **Google Cloud Console Configuration:**
    > **IMPORTANT:** This project uses **OAuth 2.0 (OAuth client ID)** for authentication. Do **NOT** use an "API key," as it will not provide the necessary permissions for calendar access.

    Follow these steps exactly to set up your API access:
    *   **Create Project:** Go to the [Google Cloud Console](https://console.cloud.google.com/), click the project dropdown (top left) > **New Project**, name it, and click **Create**.
    *   **Enable API:** Navigate to **APIs & Services > Library**, search for "Google Calendar API", click it, and then click **Enable**.
    *   **OAuth Consent Screen:**
        1.  Go to **APIs & Services > OAuth consent screen**.
        2.  Select **User Type**: **External** and click **Create**.
        3.  **App Information:** Fill in "App name", "User support email", and "Developer contact info" (your email). Click **Save and Continue**.
        4.  **Scopes:** Click **Add or Remove Scopes**. Search for `.../auth/calendar.readonly`, check the box, and click **Update**. Click **Save and Continue**.
        5.  **Test Users:** (CRITICAL) Click **+ ADD USERS**. Enter the email address of the account whose calendar you want to extract. **The script will only work for accounts added here while the app is in "Testing" mode.** Click **Save and Continue**.
    *   **Create Credentials (OAuth client ID):**
        1.  Go to **APIs & Services > Credentials**.
        2.  Click **+ CREATE CREDENTIALS** > **OAuth client ID**. (Do NOT select "API key").
        3.  Select **Application type**: **Desktop app**.
        4.  Name it (e.g., "Calendar Extractor") and click **Create**.
        5.  In the "OAuth client created" dialog, click **DOWNLOAD JSON**.
        6.  Rename this file to `credentials.json` and move it into the project folder.

2.  **Virtual Environment Setup:**
    ```bash
    python -m venv google-calendar-export
    cd google-calendar-export
    scripts\activate
    python -m pip install pip --upgrade pip
    ```

3.  **Dependencies:**
    Install required packages using:
    ```bash
    pip install -r requirements.txt
    ```

## Building and Running

### Commands
*   **Extract Events:**
    ```bash
    python main.py --days 45
    ```
*   **Custom Output:**
    ```bash
    python main.py --days 30 --output my_calendar.xlsx
    ```
*   **Switch Accounts / Force Re-authentication:**
    ```bash
    python main.py --reauth --days 7
    ```

### Arguments
*   `--days NUMBER`: (Required) Number of days in the past to look for events.
*   `--output FILENAME`: (Optional) The name of the Excel file. Defaults to `YYYY-MM-DD-Nd-calendar_export.xlsx` (e.g., `2026-04-17-45d-calendar_export.xlsx`).
*   `--reauth`: (Optional) Deletes the existing `token.pickle` and forces a new login. Use this to switch to a different Google account.

## Important Notes

*   **Whose data is pulled?** The script extracts data from the account that completes the OAuth login in the browser.
*   **Multiple Users:** You can share the `credentials.json` with others, but **every user's email address must be added to the "Test users" list** in the Google Cloud Console (unless you publish/verify the app).
*   **Authentication Token:** After the first run, a `token.pickle` file is created. This stores your login session. If you want to log in with a different account, use the `--reauth` flag or delete this file.
