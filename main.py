#!/usr/bin/env python3
"""
Google Calendar Data Extractor
Extracts calendar events from the last N days and exports to Excel spreadsheet
"""

import os
import re
import sys
from datetime import datetime, timedelta, timezone
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import argparse

try:
    from colorama import init, Fore, Style
    init(autoreset=False)  # Don't auto-reset so we can control when to reset
except ImportError:
    print("Warning: colorama not installed. Install with: pip install colorama")
    print("Continuing without colored output...\n")
    # Create dummy color constants if colorama not available
    class Fore:
        GREEN = ''
        RED = ''
        BLUE = ''
        RESET = ''
    class Style:
        RESET_ALL = ''

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def show_help():
    """Display help information"""
    print(f"\n{Fore.BLUE}Google Calendar Data Extractor{Style.RESET_ALL}")
    print(f"{Fore.BLUE}{'=' * 50}{Style.RESET_ALL}\n")
    print("This script extracts Google Calendar events and exports them to an Excel spreadsheet.")
    print("It will extract email addresses, phone numbers, and all event details.\n")
    print(f"{Fore.BLUE}Usage:{Style.RESET_ALL}")
    print("  python calendar_extractor.py --days NUMBER [--output FILENAME]\n")
    print(f"{Fore.BLUE}Arguments:{Style.RESET_ALL}")
    print("  --days NUMBER     Number of days to look back (required)")
    print("                    Example: --days 45")
    print("  --output FILENAME Output Excel filename (optional)")
    print("                    Default: calendar_export.xlsx")
    print("                    Example: --output my_calendar.xlsx\n")
    print(f"{Fore.BLUE}Examples:{Style.RESET_ALL}")
    print("  python calendar_extractor.py --days 45")
    print("  python calendar_extractor.py --days 90 --output quarterly_report.xlsx")
    print("  python calendar_extractor.py --days 7 --output last_week.xlsx\n")
    print(f"{Fore.BLUE}First-time setup:{Style.RESET_ALL}")
    print("  1. Ensure you have credentials.json in the same directory")
    print("  2. The script will open a browser for Google authorization")
    print("  3. After authorization, a token.pickle file will be created for future use\n")
    print(f"{Fore.BLUE}Required files:{Style.RESET_ALL}")
    print("  - credentials.json (Google OAuth credentials)")
    print("  - This script (calendar_extractor.py)\n")

def authenticate_google_calendar():
    """Authenticate and return Google Calendar service object"""
    creds = None
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print(f"{Fore.RED}Error: credentials.json not found!{Style.RESET_ALL}")
                print("Please ensure credentials.json is in the same directory as this script.")
                print("See README for instructions on creating credentials.json")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    service = build('calendar', 'v3', credentials=creds)
    return service

def extract_emails(text):
    """Extract email addresses from text"""
    if not text:
        return []
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(email_pattern, text)

def extract_phone_numbers(text):
    """Extract phone numbers from text"""
    if not text:
        return []
    # Pattern for various phone number formats
    phone_patterns = [
        r'\+?1?\s*\(?[0-9]{3}\)?[\s.-]?[0-9]{3}[\s.-]?[0-9]{4}',  # US format
        r'\+?[0-9]{1,3}[\s.-]?[0-9]{1,4}[\s.-]?[0-9]{1,4}[\s.-]?[0-9]{1,9}',  # International
        r'\([0-9]{3}\)\s*[0-9]{3}-[0-9]{4}',  # (xxx) xxx-xxxx
        r'[0-9]{3}-[0-9]{3}-[0-9]{4}',  # xxx-xxx-xxxx
    ]
    
    phone_numbers = []
    for pattern in phone_patterns:
        matches = re.findall(pattern, text)
        phone_numbers.extend(matches)
    
    # Clean and deduplicate
    cleaned_numbers = []
    for num in phone_numbers:
        cleaned = re.sub(r'[^\d+]', '', num)
        if len(cleaned) >= 10 and cleaned not in cleaned_numbers:
            cleaned_numbers.append(num)
    
    return cleaned_numbers

def get_calendar_events(service, days_back=45):
    """Fetch calendar events from the last N days"""
    # Calculate time range
    now = datetime.now(timezone.utc)
    time_min = now - timedelta(days=days_back)
    
    print(f"{Fore.BLUE}Fetching events from {time_min.strftime('%Y-%m-%d')} to {now.strftime('%Y-%m-%d')}...{Style.RESET_ALL}")
    
    events_result = service.events().list(
        calendarId='primary',
        timeMin=time_min.isoformat(),
        timeMax=now.isoformat(),
        singleEvents=True,
        orderBy='startTime',
        maxResults=2500  # Adjust if needed
    ).execute()
    
    events = events_result.get('items', [])
    
    # Check for additional pages
    while 'nextPageToken' in events_result:
        page_token = events_result['nextPageToken']
        events_result = service.events().list(
            calendarId='primary',
            timeMin=time_min.isoformat(),
            timeMax=now.isoformat(),
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token,
            maxResults=2500
        ).execute()
        events.extend(events_result.get('items', []))
    
    return events

def parse_event_data(events):
    """Parse event data into structured format"""
    parsed_events = []
    
    for event in events:
        # Combine all text fields for email/phone extraction
        combined_text = ' '.join(filter(None, [
            event.get('summary', ''),
            event.get('description', ''),
            event.get('location', ''),
        ]))
        
        # Extract attendees' emails
        attendees_list = event.get('attendees', [])
        attendee_emails = [att.get('email', '') for att in attendees_list]
        attendee_names = [att.get('displayName', '') for att in attendees_list]
        attendee_statuses = [att.get('responseStatus', '') for att in attendees_list]
        
        # Extract emails from text fields
        text_emails = extract_emails(combined_text)
        
        # Combine all emails
        all_emails = list(set(attendee_emails + text_emails))
        
        # Extract phone numbers
        phone_numbers = extract_phone_numbers(combined_text)
        
        # Get start and end times
        start = event.get('start', {})
        end = event.get('end', {})
        
        start_time = start.get('dateTime', start.get('date', ''))
        end_time = end.get('dateTime', end.get('date', ''))
        
        # Parse datetime strings
        try:
            if 'T' in start_time:
                start_dt = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
                end_dt = datetime.fromisoformat(end_time.replace('Z', '+00:00'))
                all_day = False
            else:
                start_dt = datetime.strptime(start_time, '%Y-%m-%d')
                end_dt = datetime.strptime(end_time, '%Y-%m-%d')
                all_day = True
        except:
            start_dt = None
            end_dt = None
            all_day = False
        
        # Calculate duration
        duration = None
        if start_dt and end_dt:
            duration = (end_dt - start_dt).total_seconds() / 3600  # in hours
        
        # Get conference data
        conference_data = event.get('conferenceData', {})
        conference_solution = conference_data.get('conferenceSolution', {}).get('name', '')
        entry_points = conference_data.get('entryPoints', [])
        meeting_links = [ep.get('uri', '') for ep in entry_points if ep.get('entryPointType') == 'video']
        
        parsed_event = {
            'event_id': event.get('id', ''),
            'summary': event.get('summary', ''),
            'description': event.get('description', ''),
            'location': event.get('location', ''),
            'start_date': start_dt.date() if start_dt else None,
            'start_time': start_dt.time() if start_dt and not all_day else None,
            'end_date': end_dt.date() if end_dt else None,
            'end_time': end_dt.time() if end_dt and not all_day else None,
            'all_day': all_day,
            'duration_hours': duration,
            'status': event.get('status', ''),
            'visibility': event.get('visibility', 'default'),
            'organizer_email': event.get('organizer', {}).get('email', ''),
            'organizer_name': event.get('organizer', {}).get('displayName', ''),
            'creator_email': event.get('creator', {}).get('email', ''),
            'creator_name': event.get('creator', {}).get('displayName', ''),
            'attendee_emails': '; '.join(attendee_emails),
            'attendee_names': '; '.join(attendee_names),
            'attendee_statuses': '; '.join(attendee_statuses),
            'attendee_count': len(attendees_list),
            'all_extracted_emails': '; '.join(all_emails),
            'extracted_phone_numbers': '; '.join(phone_numbers),
            'recurring_event_id': event.get('recurringEventId', ''),
            'is_recurring': 'recurringEventId' in event,
            'html_link': event.get('htmlLink', ''),
            'conference_type': conference_solution,
            'meeting_links': '; '.join(meeting_links),
            'created': event.get('created', ''),
            'updated': event.get('updated', ''),
            'reminders': str(event.get('reminders', {})),
            'attachments': '; '.join([att.get('title', '') for att in event.get('attachments', [])]),
            'color_id': event.get('colorId', ''),
            'transparency': event.get('transparency', 'opaque'),
        }
        
        parsed_events.append(parsed_event)
    
    return parsed_events

def export_to_excel(events_data, filename='calendar_export.xlsx'):
    """Export parsed events to Excel with formatting"""
    # Create DataFrame
    df = pd.DataFrame(events_data)
    
    # Sort by start date
    df = df.sort_values('start_date', ascending=False)
    
    # Create Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Calendar Events', index=False)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Calendar Events']
        
        # Format header row
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
        
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Add filters
        worksheet.auto_filter.ref = worksheet.dimensions
    
    print(f"{Fore.GREEN}Data exported to {filename}{Style.RESET_ALL}")
    return filename

def main():
    """Main execution function"""
    parser = argparse.ArgumentParser(description='Extract Google Calendar data to spreadsheet', add_help=False)
    parser.add_argument('--days', type=int, help='Number of days to look back')
    parser.add_argument('--output', type=str, default='calendar_export.xlsx', help='Output filename')
    parser.add_argument('--help', '-h', action='store_true', help='Show help message')
    
    args = parser.parse_args()
    
    # Show help if requested or if no days argument provided
    if args.help or not args.days:
        show_help()
        sys.exit(0 if args.help else 1)
    
    # Validate days argument
    if args.days <= 0:
        print(f"{Fore.RED}Error: --days must be a positive number{Style.RESET_ALL}")
        sys.exit(1)
    
    print(f"{Fore.BLUE}Google Calendar Data Extractor{Style.RESET_ALL}")
    print(f"{Fore.BLUE}{'=' * 50}{Style.RESET_ALL}\n")
    
    print(f"{Fore.BLUE}Authenticating with Google Calendar...{Style.RESET_ALL}")
    try:
        service = authenticate_google_calendar()
        if not service:
            print(f"{Fore.RED}Authentication failed!{Style.RESET_ALL}")
            sys.exit(1)
        print(f"{Fore.GREEN}Authentication successful!{Style.RESET_ALL}\n")
    except Exception as e:
        print(f"{Fore.RED}Authentication failed: {e}{Style.RESET_ALL}")
        print("\nMake sure you have:")
        print("1. Created a project in Google Cloud Console")
        print("2. Enabled Google Calendar API")
        print("3. Downloaded credentials.json to this directory")
        print("4. Added yourself as a test user in OAuth consent screen")
        sys.exit(1)
    
    print(f"{Fore.BLUE}Fetching events from the last {args.days} days...{Style.RESET_ALL}")
    try:
        events = get_calendar_events(service, args.days)
        print(f"{Fore.GREEN}Found {len(events)} events{Style.RESET_ALL}\n")
    except Exception as e:
        print(f"{Fore.RED}Failed to fetch events: {e}{Style.RESET_ALL}")
        sys.exit(1)
    
    if not events:
        print(f"{Fore.RED}No events found in the specified time range.{Style.RESET_ALL}")
        sys.exit(0)
    
    print(f"{Fore.BLUE}Parsing event data and extracting information...{Style.RESET_ALL}")
    try:
        parsed_events = parse_event_data(events)
    except Exception as e:
        print(f"{Fore.RED}Failed to parse events: {e}{Style.RESET_ALL}")
        sys.exit(1)
    
    # Summary statistics
    all_emails = set()
    all_phones = set()
    for event in parsed_events:
        if event['all_extracted_emails']:
            all_emails.update(event['all_extracted_emails'].split('; '))
        if event['extracted_phone_numbers']:
            all_phones.update(event['extracted_phone_numbers'].split('; '))
    
    print(f"\n{Fore.BLUE}Summary:{Style.RESET_ALL}")
    print(f"  Total events: {len(parsed_events)}")
    print(f"  Unique email addresses found: {len(all_emails)}")
    print(f"  Unique phone numbers found: {len(all_phones)}")
    
    print(f"\n{Fore.BLUE}Exporting to Excel...{Style.RESET_ALL}")
    try:
        export_to_excel(parsed_events, args.output)
    except Exception as e:
        print(f"{Fore.RED}Failed to export to Excel: {e}{Style.RESET_ALL}")
        sys.exit(1)
    
    print(f"\n{Fore.GREEN}Done! Your calendar data has been extracted to '{args.output}'{Style.RESET_ALL}")
    print(f"\n{Fore.BLUE}The spreadsheet includes:{Style.RESET_ALL}")
    print("  - Event details (title, description, location)")
    print("  - Date and time information")
    print("  - Attendee information and status")
    print("  - Extracted email addresses and phone numbers")
    print("  - Meeting links and conference details")
    print("  - Event metadata (creator, organizer, etc.)")
    
    # Reset colors before exiting
    print(Style.RESET_ALL)

if __name__ == '__main__':
    main()