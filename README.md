# Google Calendar Data Extractor - Complete Setup Guide

A Python script to extract Google Calendar events from the last 45 days (or custom period) and export them to an Excel spreadsheet with emails, phone numbers, and all event details.


![Demo](https://github.com/rubysash/google-calendar-export/blob/main/demo.png?raw=true)

## Table of Contents
- [Prerequisites](#prerequisites)
- [Step 1: Install Python](#step-1-install-python)
- [Step 2: Install Required Packages](#step-2-install-required-packages)
- [Step 3: Set Up Google Cloud Project](#step-3-set-up-google-cloud-project)
- [Step 4: Enable Google Calendar API](#step-4-enable-google-calendar-api)
- [Step 5: Create OAuth Credentials](#step-5-create-oauth-credentials)
- [Step 6: Configure OAuth Consent Screen](#step-6-configure-oauth-consent-screen)
- [Step 7: Add Test Users](#step-7-add-test-users)
- [Step 8: Download and Set Up Files](#step-8-download-and-set-up-files)
- [Step 9: Run the Script](#step-9-run-the-script)
- [Step 10: Authorize Access](#step-10-authorize-access)
- [Usage Options](#usage-options)
- [Troubleshooting](#troubleshooting)
- [What Data is Extracted](#what-data-is-extracted)
- [Sharing with Other Users](#sharing-with-other-users)
- [Cost Information](#cost-information)

## Prerequisites

- Windows 10/11 computer
- Internet connection
- Google account with Google Calendar
- Administrator access on your computer (for Python installation)

## Step 1: Install Python

### Download Python
1. Navigate to https://www.python.org/downloads/
2. Download the latest Python version (3.12 or newer)
3. Select "Windows installer (64-bit)"

### Install Python
1. Run the downloaded installer
2. **IMPORTANT**: Check the box "Add Python to PATH" at the bottom of the installer window
3. Click "Install Now"
4. Wait for installation to complete
5. Click "Close"

### Verify Installation
1. Press Windows + R, type cmd, press Enter
2. In the command prompt, type:
   ```
   python --version
   ```
3. You should see output like: Python 3.x.x
4. If you see an error, restart your computer and try again

## Step 2: Install Required Packages

1. Open Command Prompt (Windows + R, type cmd, press Enter)
2. Copy and paste this entire command:
   ```
   pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas openpyxl colorama
   ```
3. Press Enter and wait for all packages to install
4. You should see "Successfully installed..." messages

## Step 3: Set Up Google Cloud Project

### Create or Select a Project
1. Go to https://console.cloud.google.com/
2. Sign in with your Google account
3. At the top of the page, click on "Select a project" dropdown
4. Click "NEW PROJECT"
5. Enter project name: "calendar-extract"
6. Click "CREATE"
7. Wait for the project to be created (notification will appear)

## Step 4: Enable Google Calendar API

1. Ensure your "calendar-extract" project is selected in the top dropdown
2. In the search bar at the top, type "Google Calendar API"
3. Click on "Google Calendar API" from the results
4. Click the "ENABLE" button
5. Wait for the API to be enabled (page will refresh)

## Step 5: Create OAuth Credentials

### Navigate to Credentials
1. In the left sidebar, click "Credentials"
2. Click "+ CREATE CREDENTIALS" at the top
3. Select "OAuth client ID"

### Configure OAuth Client
1. If prompted to configure consent screen first, click "CONFIGURE CONSENT SCREEN"
   - If not prompted, skip to "Create OAuth Client ID" below
2. Otherwise, continue with these settings:

### Create OAuth Client ID
1. For "Application type", select **"Desktop app"**
2. For "Name", enter: "Calendar Extractor"
3. Click "CREATE"
4. A popup will appear with your client ID and secret
5. Click "OK" to close the popup

## Step 6: Configure OAuth Consent Screen

### Navigate to OAuth Consent Screen
1. In the left sidebar, click "OAuth consent screen"
2. Select "External" user type
3. Click "CREATE"

### App Information
1. App name: "calendar-extract"
2. User support email: Select your email address
3. Developer contact information: Enter your email address
4. Click "SAVE AND CONTINUE"

### Scopes Configuration
1. Click "ADD OR REMOVE SCOPES"
2. In the filter box, search for "calendar.readonly"
3. Check ONLY the box for: `.../auth/calendar.readonly`
   - Description: "See and download any calendar you can access using your Google Calendar"
4. Do NOT select any other permissions
5. Click "UPDATE"
6. Click "SAVE AND CONTINUE"

### Test Users
1. This step is handled in Step 7 below
2. Click "SAVE AND CONTINUE"

### Summary
1. Review the settings
2. Click "BACK TO DASHBOARD"

## Step 7: Add Test Users

**IMPORTANT: This step prevents the "Error 403: access_denied" error**

1. On the OAuth consent screen page, find the "Test users" section
2. Click "+ ADD USERS"
3. Enter your Gmail address
4. If sharing with another user, add their Gmail address too
5. Click "ADD"
6. Click "SAVE"

Note: You can add up to 100 test users. Only these users can run the script while in testing mode.

## Step 8: Download and Set Up Files

### Download credentials.json
1. Go to "Credentials" in the left sidebar
2. Under "OAuth 2.0 Client IDs", find "Calendar Extractor"
3. Click the download button (down arrow) on the right
4. Save the file as `credentials.json`

### Create Project Folder
1. Create a new folder on your desktop called "CalendarExtractor"
2. Move the downloaded `credentials.json` file into this folder
3. Save the Python script (from earlier) as `calendar_extractor.py` in the same folder

### Verify Folder Contents
Your CalendarExtractor folder should now contain:
- `credentials.json` (your OAuth credentials)
- `calendar_extractor.py` (the Python script)

## Step 9: Run the Script

### Open Command Prompt in Project Folder
1. Open the CalendarExtractor folder in File Explorer
2. Click in the address bar
3. Type `cmd` and press Enter
4. A Command Prompt will open in that folder

### Execute the Script
1. Type the following command:
   ```
   python calendar_extractor.py
   ```
2. Press Enter

## Step 10: Authorize Access

### First-Time Authorization
1. Your default web browser will open automatically
2. Select your Google account
3. You may see "Google hasn't verified this app" warning
4. Click "Continue" or "Advanced" then "Go to calendar-extract (unsafe)"
5. Review the permissions (should only show calendar read access)
6. Click "Continue" or "Allow"
7. You'll see "The authentication flow has completed"
8. Close the browser tab

### Return to Command Prompt
1. The script will now run and show progress:
   - "Authentication successful!"
   - "Fetching events from the last 45 days..."
   - "Found X events"
   - "Parsing event data and extracting information..."
   - Summary of extracted data
   - "Done! Your calendar data has been extracted to 'calendar_export.xlsx'"

## Usage Options

### Basic Usage
```
python calendar_extractor.py
```
Extracts last 45 days of calendar data to `calendar_export.xlsx`

### Custom Time Period
```
python calendar_extractor.py --days 60
```
Extracts last 60 days of calendar data

### Custom Output Filename
```
python calendar_extractor.py --output my_calendar_data.xlsx
```
Saves to custom filename

### Combined Options
```
python calendar_extractor.py --days 90 --output quarterly_calendar.xlsx
```

## Troubleshooting

### Error 403: access_denied
**Solution**: You haven't added yourself as a test user
1. Go back to Google Cloud Console
2. Navigate to "APIs & Services" > "OAuth consent screen"
3. Add your email under "Test users"
4. Save changes and try again

### "python is not recognized as an internal or external command"
**Solution**: Python wasn't added to PATH during installation
1. Uninstall Python
2. Reinstall and ensure "Add Python to PATH" is checked
3. Restart your computer

### "No module named 'google'"
**Solution**: Packages didn't install properly
1. Run the pip install command again:
   ```
   pip install --upgrade google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas openpyxl colorama
   ```

### Browser doesn't open for authorization
**Solution**: Firewall may be blocking it
1. Temporarily disable firewall
2. Run the script
3. After authorization, re-enable firewall

### "credentials.json not found"
**Solution**: File is missing or in wrong location
1. Ensure credentials.json is in the same folder as the Python script
2. Check the filename is exactly `credentials.json` (not credentials(1).json)

## What Data is Extracted

The Excel file will contain these columns:

### Event Information
- event_id: Unique identifier
- summary: Event title
- description: Full event description
- location: Event location
- status: confirmed/tentative/cancelled
- visibility: public/private/default

### Date and Time
- start_date: Event start date
- start_time: Event start time (if not all-day)
- end_date: Event end date
- end_time: Event end time (if not all-day)
- all_day: True/False for all-day events
- duration_hours: Calculated duration

### People and Contacts
- organizer_email: Event organizer's email
- organizer_name: Event organizer's name
- creator_email: Event creator's email
- creator_name: Event creator's name
- attendee_emails: All attendee emails (semicolon-separated)
- attendee_names: All attendee names (semicolon-separated)
- attendee_statuses: Response status for each attendee
- attendee_count: Number of attendees
- all_extracted_emails: All emails found in event text
- extracted_phone_numbers: All phone numbers found in event text

### Meeting Information
- conference_type: Type of video conference (Meet, Zoom, etc.)
- meeting_links: Video conference URLs
- html_link: Link to event in Google Calendar

### Metadata
- created: When event was created
- updated: Last modification time
- recurring_event_id: ID if part of recurring series
- is_recurring: True/False
- reminders: Reminder settings
- attachments: File attachments
- color_id: Calendar color coding
- transparency: busy/available setting

## Sharing with Other Users

If another user needs to extract their calendar:

### Option 1: Share Your Credentials (Simpler)
1. Send them:
   - The `calendar_extractor.py` script
   - Your `credentials.json` file
   - This README file
2. They follow Steps 1-2 (Python installation)
3. They run the script (Steps 9-10)
4. They authenticate with THEIR Google account
5. The script extracts THEIR calendar data

### Option 2: They Create Own Credentials (More Secure)
1. Send them:
   - The `calendar_extractor.py` script
   - This README file
2. They follow all steps 1-10
3. They create their own Google Cloud project

Note: The user only shares the final Excel file with you, never their Google credentials.

## Cost Information

### Google Calendar API Costs
- **FREE** for personal/development use
- No charges for API calls
- Free quotas:
  - 1,000,000 requests per day
  - 500 requests per 100 seconds per user

### When Costs Apply
- Never for this script's usage
- Only for large enterprise applications exceeding quotas
- Google Workspace enterprise features (not applicable here)

## Security Notes

1. **credentials.json**: Keep this file secure but it's not sensitive (it's your app's ID, not your password)
2. **token.pickle**: Created after first run, contains your actual access token - do not share this file
3. The script has read-only access - it cannot modify your calendar
4. Only approved test users can use the script while in testing mode
5. For production use (>100 users), you would need Google verification (not needed for personal use)

## Support

If you encounter issues not covered in troubleshooting:

1. Verify all steps were followed exactly
2. Check that your Google account has calendar events in the specified time period
3. Ensure you have a stable internet connection

