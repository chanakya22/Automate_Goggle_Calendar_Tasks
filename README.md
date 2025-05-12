# Google Calendar Automation

This project automates Google Calendar management using Python. It allows you to:
- Authenticate and connect to your Google Calendar.
- Download your daily schedule to a local file.
- Delete specific events (including single instances of recurring events) based on a list.

## Setup Instructions

1. **Clone the repository**

2. **Install dependencies**

   Ensure you have Python 3.7+ and pip installed. Then run:
   ```sh
   pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pytz
   ```

3. **Google API Credentials**

   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a project and enable the Google Calendar API.
   - Download your `credentials.json` and place it in the `G_Cal/` folder.
   - The first run will prompt you to authenticate and will create a `token.json` file (also in `G_Cal/`).

4. **Configure Calendar ID**

   - By default, the script uses your primary calendar. You can change the `calendar_id` variable in `Automate.py` if needed.

## Usage

You can either:
- **Download the events first** to see your daily schedule and get the correct format for deletion, or
- **Delete events directly** if you already know the required format for `delete_events.txt`.

### 1. Prepare `delete_events.txt`
- Each line should be in the format:
  ```
  <start> - <end> - <event summary>
  ```
  Example:
  ```
  2025-05-11T09:00:00-05:00 - 2025-05-11T10:00:00-05:00 - Meeting with Bob
  ```

### 2. Run the script
   ```sh
   python G_Cal/Automate.py
   ```

   - The script will:
     - Delete events listed in `delete_events.txt` for the current day (if present).
     - Download your daily schedule to `daily_schedule.txt`.

## Notes
- Sensitive files (`credentials.json`, `token.json`) are excluded from git via `.gitignore`.
- The script handles both single and recurring events. For recurring events, it attempts to cancel only the specified instance.
- All actions and errors are logged to the console for debugging.

## Troubleshooting
- If you encounter API errors, check your credentials and ensure the Calendar API is enabled.
- For recurring event instance cancellation, some Google API limitations may apply.

---

**Author:** chanuthelegend@gmail.com
