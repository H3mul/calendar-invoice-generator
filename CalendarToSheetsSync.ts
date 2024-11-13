// Script operation:
//   1. Gathers all events from matching calendars for the last month
//   2. Creates a new monthly Sheet document in the specified Drive folder
//   3. Creates a new Sheet in the document per Calendar
//   4. Populates each Sheet with summary data of the calendar's events from the last month:
//        - Date
//        - Time start/finish
//        - Duration hours in decimal format
//        - Event Name
//        - Decimal hours total
//
//
// This script is designed to fetch data from the whole previous month of the trigger time,
// expected to be triggered using a monthly time trigger (eg, 1st of every month at 12am).

function getEventDateRange(): [Date, Date] {
    // Nth day of this month
    const end = new Date();
    end.setDate(Settings.DATE_RANGE_DAY);
    end.setHours(0,0,0,0);

    // Nth day of last month
    const start = new Date(end);
    start.setDate(0); // Shift to last month
    start.setDate(Settings.DATE_RANGE_DAY);

    return [start, end];
}

type CalendarData = {
    calendar: GoogleAppsScript.Calendar.Calendar;
    events?: GoogleAppsScript.Calendar.CalendarEvent[];
}

function fetchCalendarData(): Map<string, CalendarData> {
    const invoiceData = new Map<string, CalendarData>();
    const calendars = CalendarApp.getAllCalendars()
                                 .filter(c => Settings.CALENDAR_NAME_FILTER.test(c.getName()));

    calendars.forEach(calendar => {
        invoiceData.set(
            calendar.getId(),
            { calendar, events: calendar.getEvents(...getEventDateRange()) }
        );
    });

    console.log(`Fetched last month's events for ${calendars.length} calendars`)
    return invoiceData;
}

function getFolderId() {
    return PropertiesService.getScriptProperties().getProperty("MONTHLY_SHEET_FOLDER_ID") ||
            Settings.MONTHLY_SHEET_FOLDER_ID;
}

function getPath(file: GoogleAppsScript.Drive.File) {
    let path = file.getName();
    let folder = file.getParents();
    while(folder.hasNext()) {
        path = `${folder.next()}/${path}`
    }
    return path;
}

function createMonthlyInvoiceSheet() {
    const dateRange = getEventDateRange();
    const fileName = `${Strings.sheetFileNamePrefix} ${dateRange[0].toLocaleDateString()} - ${dateRange[1].toLocaleDateString()}`;

    const folder = DriveApp.getFolderById(getFolderId());
    const sheet = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(sheet.getId());
    file.moveTo(folder);

    console.log(`Created a new monthly sheet: ${getPath(file)}`)

    return sheet
}

function formatTimeString(dateSomething: any) {
    return new Date(dateSomething).toLocaleTimeString(Settings.TIME_LOCALE, Settings.TIME_FORMAT);
}

function generateMonthlyInvoiceSheet(invoiceData: Map<string, CalendarData>) {
    const spreadsheet = createMonthlyInvoiceSheet();
    const notesSheet = spreadsheet.getActiveSheet();

    // Add a readme sheet
    notesSheet.setName(Strings.notesSheetTitle);
    notesSheet.getRange("A1").setValue(Strings.sheetGenerationInfo);
    notesSheet.getRange("A2").setValue(new Date().toString());

    const headings = [
        Strings.headings.day,
        Strings.headings.time,
        Strings.headings.title,
        Strings.headings.hours,
        Strings.headings.totalInclusion
    ];

    invoiceData.values().forEach(d => {
        const data:Array<Array<any>> = [headings];

        d.events
            .filter(e => !e.isAllDayEvent())
            .forEach(e => {
                data.push([
                    e.getStartTime().toLocaleDateString(),
                    `${formatTimeString(e.getStartTime())}-${formatTimeString(e.getEndTime())}`,
                    `${e.getTitle()}`,
                    `${(e.getEndTime().getTime() - e.getStartTime().getTime()) / (1000 * 60 * 60)}`,
                    "TRUE"
                ]);
            });

        // Append hours total row
        data.push(["", "", "Total: ", `=SUM(ARRAYFORMULA(D2:D${data.length}*E2:E${data.length}))`, ""]);

        // Create a new sheet for the calendar
        const sheet = spreadsheet.insertSheet(d.calendar.getName());

        // all the calendar data 
        sheet.getRange(1,1, data.length, data[0].length).setValues(data);

        // Add checkboxes to the inclusion column
        sheet.getRange(2, headings.indexOf(Strings.headings.totalInclusion) + 1, data.length-2, 1)
             .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

        // Set first and last rows to bold
        const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
        sheet.getRange(1,1, 1, data[0].length).setTextStyle(boldStyle);
        sheet.getRange(data.length,1, 1, data[0].length).setTextStyle(boldStyle);

        sheet.setColumnWidths(1, headings.length, Settings.DEFAULT_COLUMN_WIDTH);
        // Make Event Title column a bit wider
        sheet.setColumnWidth(headings.indexOf(Strings.headings.title) + 1, Settings.TITLE_COLUMN_WIDTH);
    });

    console.log("Finished populating monthly sheet with calendar data")
}

function onTrigger() {
    console.log(`Starting monthly-calendar-summary-script with properties: ${JSON.stringify(PropertiesService.getScriptProperties().getProperties())}`);
    const invoiceData = fetchCalendarData();
    generateMonthlyInvoiceSheet(invoiceData);
}