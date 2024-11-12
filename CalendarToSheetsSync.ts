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

    const calendar = calendars[0];

    // calendars.forEach(calendar => {
        invoiceData.set(
            calendar.getId(),
            { calendar, events: calendar.getEvents(...dateRange) }
        );

        console.log(`Calendar: ${calendar.getName()}`);
        console.log(JSON.stringify(
            invoiceData.get(calendar.getId()).events.map(e => e.getTitle())
        , null, 2));

    // });
    return invoiceData;
}

function onTrigger() {
    const invoiceData = fetchCalendarData();
}