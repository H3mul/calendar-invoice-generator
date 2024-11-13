// Script Source:
// https://github.com/H3mul/calendar-invoice-generator


// Script operation:
//   1. Gathers all events from matching calendars for the last month
//   2. Creates a new monthly Sheet document in the specified Drive folder
//   3. Creates a new Sheet in the document per Calendar
//   4. Populates each Sheet with summary data of the calendar's events from the last month
//
//
// This script is designed to fetch data from the whole previous month of the trigger time,
// onTrigger() is expected to be triggered using a monthly time trigger (eg, 1st of every month at 12am).

// Script Properties:
//    MONTHLY_SHEET_FOLDER_ID - Optional Drive target folder Id (default: the sheet is created in the root folder)
//    CALENDAR_NAME_FILTER    - Optional Calendar name filter regex (default: no filtering)

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

function getFolderByIdSafe(folderId: string):GoogleAppsScript.Drive.Folder {
    if (!folderId) {
        return null
    }

    try {
        return DriveApp.getFolderById(folderId);
    } catch {
        console.log("Folder Id not found, using default");
        return null;
    }
}

function getFolder():GoogleAppsScript.Drive.Folder {
    return getFolderByIdSafe(PropertiesService.getScriptProperties().getProperty("MONTHLY_SHEET_FOLDER_ID")) ||
           DriveApp.getRootFolder();
}

function getCalendarFilterRe() {
    return new RegExp(PropertiesService.getScriptProperties().getProperty("CALENDAR_NAME_FILTER") || ".*", "i");
}

type CalendarData = {
    calendar: GoogleAppsScript.Calendar.Calendar;
    events?: GoogleAppsScript.Calendar.CalendarEvent[];
}

function fetchCalendarData(): Map<string, CalendarData> {
    const invoiceData = new Map<string, CalendarData>();
    const calendars = CalendarApp.getAllCalendars()
                                 .filter(c => getCalendarFilterRe().test(c.getName()));

    calendars.forEach(calendar => {
        invoiceData.set(
            calendar.getId(),
            { calendar, events: calendar.getEvents(...getEventDateRange()) }
        );
    });

    console.log(`Fetched last month's events for ${calendars.length} calendars`)
    return invoiceData;
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

    const sheet = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(sheet.getId());
    file.moveTo(getFolder());

    console.log(`Created a new monthly sheet: ${getPath(file)}`)

    return sheet
}

function formatTimeString(dateSomething: any) {
    return new Date(dateSomething).toLocaleTimeString(Settings.TIME_LOCALE, Settings.TIME_FORMAT);
}

function generateMonthlyInvoiceSheet(invoiceData: Map<string, CalendarData>) {
    const spreadsheet = createMonthlyInvoiceSheet();
    const aboutSheet = spreadsheet.getActiveSheet();

    // Add a readme sheet
    aboutSheet.setName(Strings.aboutSheetTitle);
    aboutSheet.getRange("A1").setValue(Strings.sheetGenerationInfo);
    aboutSheet.getRange("A2").setValue(new Date().toString());

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

    // Move the about sheet to the back and select the first sheet as active
    spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
    spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);

    console.log("Finished populating monthly sheet with calendar data")
}

function onTrigger() {
    console.log(`Starting monthly-calendar-summary script with properties: ${JSON.stringify(PropertiesService.getScriptProperties().getProperties())}`);
    const invoiceData = fetchCalendarData();
    generateMonthlyInvoiceSheet(invoiceData);
}