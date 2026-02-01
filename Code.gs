/***********************
 MANUAL CONFIGURATION(S)
***********************/
const CONFIG = {

  // Align Sheet Tabs and Subscriptable Calendars.
  sheets: {

    // Sample Template:
    // SheetTab     : "CalendarName",
    PhotoCompEvents : "PhotoComps",
    HackCompEvents  : "Hackathons",
  },

  // Ensure total entries of "maxRows" rows by duplicating entry at row "templateRow".
  maxRows: 1000,
  templateRow: 900,

  // Clean-up entry from Google Sheets after "cleanupDays" days.
  // Note: Event will NOT be deleted from Google Calendar after clean-up.
  cleanupDays: 3,

  // If "assumeEventStartsOnSync" is true, event startDate is set to the day it was synced;
  // If "assumeEventStartsOnSync" is true, event startDate is set to assumeEventLengthDays days before endDate.
  assumeEventLengthDays: 2,
  assumeEventStartsOnSync: false,
};

/***********************
 MAIN ENTRY
***********************/
function syncAllEventSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName in CONFIG.sheets) {
    const calendarName = CONFIG.sheets[sheetName];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const calendar = getCalendarByName(calendarName);
    if (!calendar) throw new Error(`Calendar not found: ${calendarName}`);

    cleanupPastEvents(sheet);
    syncSheetToCalendar(sheet, calendar);
    sortByNearestDate(sheet);
    enforceRowCount(sheet);
  }
}

/***********************
 CORE SYNC LOGIC
***********************/
function syncSheetToCalendar(sheet, calendar) {
  const data = sheet.getDataRange().getValues();
  const today = stripTime(new Date());

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];

    const [
      website,   // A
      name,      // B
      startDate, // C
      endDate,   // D
      url,       // E
      ,          // F LastUpdated (Column F)
      eventId,   // G
      status,    // H
      priority,  // I
      delFlag,   // J
    ] = row;

    /* =========================
       DELETE HANDLING (COLUMN J)
       ========================= */
    if (String(delFlag).toUpperCase() === 'YES') {
      if (eventId) {
        try {
          const event = calendar.getEventById(eventId);
          if (event) event.deleteEvent();
        } catch (e) {}
      }
      sheet.deleteRow(i + 1);
      continue;
    } else if (delFlag !== '') {
      sheet.getRange(i + 1, 10).setValue('');
    }

    if (!website || !name || !endDate) continue;

    const end = stripTime(new Date(endDate));
    if (end < today) continue;

    let start = startDate
      ? stripTime(new Date(startDate))
      : null;

    const title = `[${website}] ${name}`;
    const description = url ? `Source: ${url}` : '';
    let event = null;
    let isNewEvent = false;

    /* =========================
       UPDATE EXISTING EVENT
       ========================= */
    if (eventId) {
      try {
        event = calendar.getEventById(eventId);
        if (event) {
          if (!start) {
            start = new Date(
              end.getTime() - CONFIG.assumeEventLengthDays * 86400000
            );
          }

          event.setTitle(title);
          event.setDescription(description);
          event.setAllDayDates(start, new Date(end.getTime() + 86400000));
        }
      } catch (e) {
        event = null;
      }
    }

    /* =========================
       CREATE NEW EVENT
       ========================= */
    if (!event) {
      if (!start) {
        start = CONFIG.assumeEventStartsOnSync
          ? today
          : new Date(
              end.getTime() -
              CONFIG.assumeEventLengthDays * 86400000
            );

        sheet.getRange(i + 1, 3).setValue(start);
      }

      event = calendar.createAllDayEvent(
        title,
        start,
        new Date(end.getTime() + 86400000),
        { description }
      );

      sheet.getRange(i + 1, 7).setValue(event.getId());
      isNewEvent = true;
    }

    /* =========================
       PRIORITY HANDLING (COLUMN I)
       ========================= */
    if (priority === true || String(priority).toUpperCase() === 'TRUE') {
      event.setColor(CalendarApp.EventColor.RED);
      sheet.getRange(i + 1, 9).setValue('TRUE');
    } else {
      sheet.getRange(i + 1, 9).setValue('');
    }

    /* =========================
       TIMESTAMPS (COLUMN F) AND STATUS (COLUMN H)
       ========================= */
    sheet.getRange(i + 1, 6).setValue(new Date()); // LastUpdated (Column F)
    sheet.getRange(i + 1, 8).setValue(
      isNewEvent ? 'SYNCED' : 'UPDATED'
    );
  }
}

/***********************
 CLEANUP LOGIC - ENSURE ONLY PRESENT EVENTS ARE DISPLAYED;
 IF OTHERWISE: DELETE ALL PAST EVENTS cleanupDays AFTER ENDED.
 NOTE: ALL PREVIOUS/EXISTING EVENTS STILL REMAIN STORED IN GOOGLE CALENDAR.
***********************/
function cleanupPastEvents(sheet) {
  const data = sheet.getDataRange().getValues();
  const cutoff = stripTime(
    new Date(Date.now() - CONFIG.cleanupDays * 86400000)
  );

  for (let i = data.length - 1; i >= 1; i--) {
    const endDate = data[i][3];
    if (!endDate) continue;

    const end = stripTime(new Date(endDate));
    if (end < cutoff) {
      sheet.deleteRow(i + 1);
    }
  }
}

/***********************
 SORT BY NEAREST DATE - WHICH EVENT COMES FIRST?
***********************/
function sortByNearestDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({ column: 4, ascending: true });
}

/***********************
 ROW COUNT ENFORCER - ENSURE ROW COUNT ALWAYS maxRows;
 OTHERWISE: DUPLICATE templateRow UNTIL NUMBER OF ROWS = maxRows.
***********************/
function enforceRowCount(sheet) {
  const currentRows = sheet.getLastRow();
  if (currentRows >= CONFIG.maxRows) return;

  const templateRow = CONFIG.templateRow;
  if (templateRow > sheet.getLastRow()) return;

  const template = sheet
    .getRange(templateRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const rowsToAdd = CONFIG.maxRows - currentRows;
  for (let i = 0; i < rowsToAdd; i++) {
    sheet.appendRow(template);
  }
}

/***********************
 HELPERS
***********************/
function getCalendarByName(name) {
  return CalendarApp.getCalendarsByName(name)[0] || null;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}
