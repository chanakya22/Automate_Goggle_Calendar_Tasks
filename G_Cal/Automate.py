import os
import sys
from datetime import datetime, timedelta
import pytz

# Ensure required packages are installed
try:
    from googleapiclient.discovery import build
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("Required packages not found. Installing...")
    os.system(f"{sys.executable} -m pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib")
    from googleapiclient.discovery import build
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

# Authenticate and build the service
SCOPES = ['https://www.googleapis.com/auth/calendar']
CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), 'credentials.json')
TOKEN_PATH = os.path.join(os.path.dirname(__file__), 'token.json')

# Check if token.json exists
if os.path.exists(TOKEN_PATH):
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
else:
    # Run the OAuth flow to generate token.json
    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
    creds = flow.run_local_server(port=0)
    # Save the credentials for future use
    with open(TOKEN_PATH, 'w') as token:
        token.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds)

def delete_events_from_file(calendar_id, file_path):
    try:
        print(f"Reading events from file: {file_path}")
        with open(file_path, 'r') as file:
            events = file.readlines()

        print(f"Total events found in file: {len(events)}")
        for index, event in enumerate(events):
            print(f"Processing event {index + 1}: {event.strip()}")
            event_details = event.strip().split(' - ')
            if len(event_details) >= 3:
                time_min = event_details[0].strip()
                time_max = event_details[1].strip()
                event_summary = ' - '.join(event_details[2:]).strip()
                print(f"Using timeMin: {time_min}, timeMax: {time_max}, and eventSummary: {event_summary}")

                print(f"API Call: service.events().list(calendarId={calendar_id}, timeMin={time_min}, timeMax={time_max}, q={event_summary}, singleEvents=True)")
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    q=event_summary,
                    singleEvents=True
                ).execute()

                items = events_result.get('items', [])
                print(f"Found {len(items)} matching events in Google Calendar")

                for item in items:
                    print(f"Event found: {item.get('summary')} with ID: {item.get('id')}")
                    if item['summary'] == event_summary:
                        if 'recurringEventId' in item and 'originalStartTime' in item:
                            recurring_id = item['recurringEventId']
                            original_start = item['originalStartTime']
                            original_end = item['end'] if 'end' in item else None
                            print(f"Cancelling single instance of recurring event: {event_summary} (recurringEventId={recurring_id}, originalStartTime={original_start})")
                            resource = {
                                'recurringEventId': recurring_id,
                                'originalStartTime': {},
                                'status': 'cancelled'
                            }
                            if 'dateTime' in original_start:
                                resource['originalStartTime']['dateTime'] = original_start['dateTime']
                                if 'timeZone' in original_start:
                                    resource['originalStartTime']['timeZone'] = original_start['timeZone']
                                if original_end and 'dateTime' in original_end:
                                    resource['start'] = {'dateTime': original_start['dateTime']}
                                    resource['end'] = {'dateTime': original_end['dateTime']}
                                    if 'timeZone' in original_end:
                                        resource['end']['timeZone'] = original_end['timeZone']
                            elif 'date' in original_start:
                                resource['originalStartTime']['date'] = original_start['date']
                                if original_end and 'date' in original_end:
                                    resource['start'] = {'date': original_start['date']}
                                    resource['end'] = {'date': original_end['date']}
                            print(f"API Call: service.events().insert(calendarId={calendar_id}, body={resource})")
                            service.events().insert(
                                calendarId=calendar_id,
                                body=resource
                            ).execute()
                            print(f"Successfully cancelled single instance of recurring event: {event_summary}")
                        else:
                            print(f"Deleting event: {event_summary} from {time_min} to {time_max}")
                            service.events().delete(
                                calendarId=calendar_id,
                                eventId=item['id'],
                                sendUpdates='none'
                            ).execute()
                            print(f"Successfully deleted event: {event_summary}")
                    else:
                        print(f"Event summary does not match: {item.get('summary')} != {event_summary}")
            else:
                print(f"Skipping invalid event format: {event.strip()}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Define the output file path in the G_Cal folder
OUTPUT_FILE_PATH = os.path.join(os.path.dirname(__file__), 'daily_schedule.txt')

def download_daily_schedule(calendar_id, date):
    try:
        # Check if the schedule has already been downloaded for the day
        # if os.path.exists(OUTPUT_FILE_PATH):
        #     with open(OUTPUT_FILE_PATH, 'r') as file:
        #         first_line = file.readline().strip()
        #         if first_line.startswith(f"# Schedule for {date}"):
        #             print(f"Schedule for {date} is already downloaded. Skipping download.")
        #             return

        # Write a header to indicate the schedule date
        with open(OUTPUT_FILE_PATH, 'w') as file:
            file.write(f"# Schedule for {date}\n")

        # Convert the date to the start and end of the day in local timezone
        local_tz = pytz.timezone('America/Chicago')  # Replace with your timezone
        start_of_day = local_tz.localize(datetime.strptime(date, '%Y-%m-%d'))
        end_of_day = start_of_day + timedelta(days=1) - timedelta(seconds=1)

        # Get the current hour
        current_hour = datetime.now().hour

        # Loop through each hour of the day up to the current hour
        for hour in range(current_hour+1):
            start_of_hour = start_of_day + timedelta(hours=hour)
            end_of_hour = start_of_hour + timedelta(hours=1) - timedelta(seconds=1)

            # Convert to UTC for the API
            time_min = start_of_hour.astimezone(pytz.utc).isoformat()
            time_max = end_of_hour.astimezone(pytz.utc).isoformat()

            # print(f"Fetching events for hour: {hour:02d}:00 to {hour:02d}:59")

            # Fetch events for the specified hour
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            events = events_result.get('items', [])

            # Append events to the output file
            with open(OUTPUT_FILE_PATH, 'a') as file:
                if not events:
                    # file.write(f"No events found for hour {hour:02d}:00 to {hour:02d}:59.\n")
                    continue
                else:
                    for event in events:
                        start = event['start'].get('dateTime', event['start'].get('date'))
                        end = event['end'].get('dateTime', event['end'].get('date'))
                        file.write(f"{start} - {end} - {event['summary']}\n")

        # print(f"Daily schedule saved to {OUTPUT_FILE_PATH}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    calendar_id = 'chanuthelegend@gmail.com'  # Use 'primary' for the main calendar

    # Get the current date
    date = datetime.now().strftime('%Y-%m-%d')
    # print(f"Using current date: {date}")

    # calendar_list = service.calendarList().list().execute()
    # calendars = calendar_list.get('items', [])
    # for calendar in calendars:
    #     print(f"ID: {calendar['id']}, Name: {calendar['summary']}")


    # Delete events from file
    delete_events_file_path = os.path.join(os.path.dirname(__file__), 'delete_events.txt')
    delete_events_from_file(calendar_id, delete_events_file_path)

    # Download daily schedule
    download_daily_schedule(calendar_id, date)