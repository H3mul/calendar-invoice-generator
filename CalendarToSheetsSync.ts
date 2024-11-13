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

import {settings} from './Settings';

function getEventDateRange(): [Date, Date] {
    // Nth day of this month
    const end = new Date();
    end.setDate(settings.DATE_RANGE_DAY);
    end.setHours(0,0,0,0);

    // Nth day of last month
    const start = new Date(end);
    start.setDate(0); // Shift to last month
    start.setDate(settings.DATE_RANGE_DAY);

    return [start, end];
}

type CalendarData = {
    calendar: GoogleAppsScript.Calendar.Calendar;
    events?: GoogleAppsScript.Calendar.CalendarEvent[];
}

function fetchCalendarData(): Map<string, CalendarData> {
    const invoiceData = new Map<string, CalendarData>();

    const calendars = CalendarApp.getAllCalendars()
                                 .filter(c => settings.CALENDAR_NAME_FILTER.test(c.getName()));
    
    const dateRange = getEventDateRange();


    // DEBUG: temporarily only use first calendar instead of all
    // const calendar = calendars[0];
    calendars.forEach(calendar => {
        invoiceData.set(
            calendar.getId(),
            { calendar, events: calendar.getEvents(...dateRange) }
        );
    });
    return invoiceData;
}

function createMonthlyInvoiceSheet() {
    const dateRange = getEventDateRange();
    const fileName = `Monthly Calendar Summary ${dateRange[0].toLocaleDateString()} - ${dateRange[1].toLocaleDateString()}`;

    const folder = DriveApp.getFolderById(settings.MONTHLY_SHEET_FOLDER_ID);
    const sheet = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(sheet.getId());
    return sheet
}

function formatTimeString(dateSomething: any) {
    return new Date(dateSomething).toLocaleTimeString(settings.TIME_LOCALE, settings.TIME_FORMAT);
}

function generateMonthlyInvoiceSheet(invoiceData: Map<string, CalendarData>) {
    // DEBUG: temporarily grab test sheet instead of creating new one
    // const spreadsheet = createMonthlyInvoiceSheet();
    const spreadsheet = SpreadsheetApp.openById(settings.MONTHLY_SHEET_TEST_FILE_ID);

    // DEBUG: temporarily delete all other sheets
    spreadsheet.getSheets()
        .filter(s => s.getName() !== "Notes")
        .forEach(s => spreadsheet.deleteSheet(s));

    const notesSheet = spreadsheet.getActiveSheet();

    // DEBUG: temporarily clear sheet
    notesSheet.clear();
    notesSheet.setName("Notes");
    notesSheet.getRange("A1").setValue(`This Sheet has been created using a periodic App Script`)
    notesSheet.getRange("A2").setValue(`on ${new Date().toString()}`);

    const headings = [ "Day", "Timestamps", "Event Title", "Hours Count", "Include In Total" ];

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


        const sheet = spreadsheet.insertSheet(d.calendar.getName());

        // Add data
        sheet.getRange(1,1, data.length, data[0].length).setValues(data);

        // Add validation to the inclusion column (note: assume its the last column)
        sheet.getRange(2, headings.length, data.length-2, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

        // Set first and last rows to bold
        const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
        sheet.getRange(1,1, 1, data[0].length).setTextStyle(boldStyle);
        sheet.getRange(data.length,1, 1, data[0].length).setTextStyle(boldStyle);

        sheet.setColumnWidths(1, headings.length, 100);
        // Make Event Title column a bit wider
        sheet.setColumnWidth(3, 175);
    });
}

function onTrigger() {
    const invoiceData = fetchCalendarData();
    generateMonthlyInvoiceSheet(invoiceData);
}