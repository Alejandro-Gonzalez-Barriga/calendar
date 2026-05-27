// Constants
const UMBRAGE_KEYWORD = "Umbrage";
const PROJECT_CALL_TITLES = ["Project call", "Resla", "CheckSammy", "RevStar"];
const TIMEZONE = "America/Mexico_City";
const SCRIPT_TAG = "Created by Script";

// Safely normalize titles
const normalizeTitle = (title = "") => {
  const str = String(title);
  return str.startsWith(UMBRAGE_KEYWORD) ? UMBRAGE_KEYWORD : str;
};

function createCalendarEvents() {
  Logger.log("🔄 Starting calendar sync");

  const ss = SpreadsheetApp.openById("1AMGUOTTL3cVrhNcy55xRIDFllFoBAgL5iRkjvWXAa50");
  const sheet = ss.getSheetByName("Nov 2025");
  if (!sheet) {
    Logger.log("⚠️ Sheet 'Nov 2025' not found. Aborting.");
    return;
  }

  // Load calendars
  const allCals = CalendarApp.getAllCalendars();
  const projectCallCal = allCals.find(c => c.getName() === "Project call");
  const interviewsCal = allCals.find(c => c.getName() === "Interviews");

  // ✅ Avoid optional chaining here
  const projectCallName = projectCallCal ? projectCallCal.getName() : "";
  const interviewsName = interviewsCal ? interviewsCal.getName() : "";
  Logger.log(`📅 Using calendars: Project call='${projectCallName}', Interviews='${interviewsName}'`);

  // Sheet layout constants
  const startRow = 12, lastRow = 35;
  const dateRow = 3, dateCol = 2, numDateCols = 28;
  const timeCol = 1;

  // Bulk fetch sheet data
  const dateHeaders = sheet.getRange(dateRow, dateCol, 1, numDateCols).getDisplayValues()[0];
  const timeValues = sheet.getRange(startRow, timeCol, lastRow - startRow + 1).getValues().flat();
  const matrix = sheet.getRange(startRow, dateCol, lastRow - startRow + 1, numDateCols).getValues();

  // Compute base dates (midnight for each column)
  const midnights = dateHeaders.map(hdr => {
    const d = Utilities.parseDate(hdr, TIMEZONE, "MMM d, yyyy");
    return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  });

  // Build sheet event key sets
  const sheetProjKeys = new Set();
  const sheetIntKeys = new Set();
  midnights.forEach((base, colIdx) => {
    if (!base) return;
    matrix.forEach((rowData, rowIdx) => {
      const raw = (rowData[colIdx] || "").toString().trim();
      if (!raw) return;
      const title = normalizeTitle(raw);
      const tm = timeValues[rowIdx].toString().match(/(\d{1,2}):(\d{2})/);
      if (!tm) return;
      let [h, m] = tm.slice(1).map(Number);
      if (h < 8) h += 12;
      const span = getMergedRowSpan(sheet, startRow + rowIdx, dateCol + colIdx) || 1;
      const startDT = new Date(base); startDT.setHours(h, m);
      const endDT = new Date(startDT); endDT.setMinutes(endDT.getMinutes() + span * 30);
      const key = `${title}|${startDT.getTime()}|${endDT.getTime()}`;
      const isProj = PROJECT_CALL_TITLES.includes(raw) || raw.includes(UMBRAGE_KEYWORD);
      (isProj ? sheetProjKeys : sheetIntKeys).add(key);
    });
  });
  Logger.log(`🔑 Sheet keys: ${sheetProjKeys.size} project, ${sheetIntKeys.size} interview`);

  // Delete stale events
  deleteRemovedEvents(projectCallCal, sheetProjKeys, 'Project call');
  deleteRemovedEvents(interviewsCal, sheetIntKeys, 'Interviews');

  // Iterate sheet and create or skip events
  midnights.forEach((base, colIdx) => {
    if (!base) return;
    matrix.reduce((nextRow, rowData, rowIdx) => {
      if (rowIdx < nextRow) return nextRow;
      const raw = (rowData[colIdx] || "").toString().trim();
      if (!raw) return nextRow + 1;
      const title = normalizeTitle(raw);
      const tm = timeValues[rowIdx].toString().match(/(\d{1,2}):(\d{2})/);
      if (!tm) return nextRow + 1;
      let [h, m] = tm.slice(1).map(Number);
      if (h < 8) h += 12;
      const span = getMergedRowSpan(sheet, startRow + rowIdx, dateCol + colIdx) || 1;
      const startDT = new Date(base); startDT.setHours(h, m);
      const endDT = new Date(startDT); endDT.setMinutes(endDT.getMinutes() + span * 30);

      const cal = (PROJECT_CALL_TITLES.includes(raw) || raw.includes(UMBRAGE_KEYWORD))
        ? projectCallCal : interviewsCal;

      if (!cal) {
        Logger.log(`⚠️ Skipped: no calendar found for ${title} @ ${startDT}`);
        return nextRow + span;
      }

      const exists = cal.getEventsForDay(startDT).some(ev =>
        normalizeTitle(ev.getTitle()) === title &&
        ev.getStartTime().getTime() === startDT.getTime() &&
        ev.getEndTime().getTime() === endDT.getTime()
      );
      if (exists) {
        Logger.log(`⏩ Skipped: [${cal.getName()}] ${title} @ ${startDT}`);
      } else {
        cal.createEvent(title, startDT, endDT, {description: SCRIPT_TAG, timeZone: TIMEZONE});
        Logger.log(`✅ Created: [${cal.getName()}] ${title} | ${startDT}→${endDT}`);
      }
      return nextRow + span;
    }, 0);
  });
}

function deleteRemovedEvents(calendar, sheetKeys, name) {
  if (!calendar) return;
  Logger.log(`🧹 Cleaning stale for [${name}]`);
  const now = new Date();
  const future = new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());
  calendar.getEvents(now, future).forEach(ev => {
    const key = `${normalizeTitle(ev.getTitle())}|${ev.getStartTime().getTime()}|${ev.getEndTime().getTime()}`;
    if (!sheetKeys.has(key)) {
      ev.deleteEvent();
      Logger.log(`🗑️ Deleted stale: [${name}] ${ev.getTitle()} @ ${ev.getStartTime()}`);
    }
  });
}

function getMergedRowSpan(sheet, row, col) {
  const r = sheet.getRange(row, col);
  if (!r.isPartOfMerge()) return 1;
  const mr = r.getMergedRanges().find(m => m.getRow() <= row && row <= m.getLastRow());
  return mr ? mr.getNumRows() : 1;
}