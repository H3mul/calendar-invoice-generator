namespace Settings {
    export const DATE_RANGE_DAY =  1; // nth of the month
    export const TIME_LOCALE = 'en-CA';
    export const TIME_FORMAT = { hour12: false, hour: "2-digit", minute: "2-digit" } as Intl.DateTimeFormatOptions

    export const DEFAULT_COLUMN_WIDTH = 100;
    export const TITLE_COLUMN_WIDTH = 175;

    // Do not change, or regeneration of old sheets won't work
    export const sheetGenerationTimeRangeName = "GatheredCalendarDataDateRange";
}