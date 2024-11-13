# Monthly Calendar Events To Google Sheets Summary App Script

This script is intended to run as an App Script in Google Drive, and generates a Spreadsheet with a breakdown of event hours on a per-calendar basis.

## Deployment:

https://developers.google.com/apps-script/guides/clasp#create_a_new_apps_script_project
https://github.com/google/clasp/blob/master/docs/typescript.md

1. Set up clasp as in above
2. Deploy into new script in google drive using `clasp push`
3. Set up desired Project Trigger
4. Set desired Project Script Properties

## Script operation:

1. Gathers all events from matching calendars for the last month
2. Creates a new monthly Sheet document in the specified Drive folder
3. Creates a new Sheet in the document per Calendar
4. Populates each Sheet with summary data of the calendar's events from the last month

This script is designed to fetch data from the whole previous month of the trigger time,
onTrigger() is expected to be triggered using a monthly time trigger (eg, 1st of every month at 12am).

## AppScript Project Script Properties:
    MONTHLY_SHEET_FOLDER_ID - Optional Drive target folder Id (default: the sheet is created in the root folder)
    CALENDAR_NAME_FILTER    - Optional Calendar name filter regex (default: no filtering)