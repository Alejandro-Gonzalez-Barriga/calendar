// Constants
const UMBRAGE_KEYWORD = "Umbrage";
const PROJECT_CALL_TITLES = ["Project call", "Resla", "CheckSammy", "RevStar"];
const TIMEZONE = "America/Mexico_City";
const SCRIPT_TAG = "Created by Script";

const SHEET_CONFIG = Object.freeze({
  spreadsheetId: "1AMGUOTTL3cVrhNcy55xRIDFllFoBAgL5iRkjvWXAa50",
  sheetName: "Jan 2026",
  startRow: 12,
  lastRow: 35,
  dateRow: 3,
  dateCol: 3,      // column C
  numDateCols: 28,
  timeCol: 1,      // column A
});

// Safely normalize titles
const normalizeTitle = (title = "") => {
  const str = String(title || "").trim();
  return str.startsWith(UMBRAGE_KEYWORD) ? UMBRAGE_KEYWORD : str;
};

const buildEventKey = (title, startTime, endTime) =>
  `${title}|${startTime.getTime()}|${endTime.getTime()}`;

function createCalendarEvents() {
  try {
    Logger.log("🔄 Starting calendar sync");

    const sheet = getTargetSheet();
    if (!sheet) return;

    const { projectCallCal, interviewsCal } = getTargetCalendars();
    const { events, projectKeys, interviewKeys } = extractSheetEvents(sheet);

    const deletedCount =
      deleteRemovedEvents(projectCallCal, projectKeys, "Project call") +
      deleteRemovedEvents(interviewsCal, interviewKeys, "Interviews");
    Logger.log(`📋 Total deleted: ${deletedCount}`);

    const { createdCount, skippedCount } = syncSheetEvents(
      events,
      projectCallCal,
      interviewsCal
    );

    Logger.log("=".repeat(50));
    Logger.log(
      `📊 SUMMARY: Created=${createdCount}, Skipped=${skippedCount}, Deleted=${deletedCount}`
    );
    Logger.log("✅ Calendar sync completed");
  } catch (error) {
    Logger.log(`❌ ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SHEET_CONFIG.spreadsheetId);
  Logger.log("✅ Spreadsheet opened");

  const sheet = ss.getSheetByName(SHEET_CONFIG.sheetName);
  if (!sheet) {
    Logger.log(`⚠️ Sheet '${SHEET_CONFIG.sheetName}' not found. Aborting.`);
    return null;
  }
  Logger.log("✅ Sheet found");
  return sheet;
}

function getTargetCalendars() {
  const allCals = CalendarApp.getAllCalendars();
  const projectCallCal = allCals.find((c) => c.getName() === "Project call");
  const interviewsCal = allCals.find((c) => c.getName() === "Interviews");

  Logger.log(
    `📅 Using calendars: Project call='${projectCallCal ? projectCallCal.getName() : "NOT FOUND"}', ` +
      `Interviews='${interviewsCal ? interviewsCal.getName() : "NOT FOUND"}'`
  );

  return { projectCallCal, interviewsCal };
}

function extractSheetEvents(sheet) {
  const { startRow, lastRow, dateRow, dateCol, numDateCols, timeCol } =
    SHEET_CONFIG;

  const dateHeaders = sheet
    .getRange(dateRow, dateCol, 1, numDateCols)
    .getDisplayValues()[0];

  const timeValues = sheet
    .getRange(startRow, timeCol, lastRow - startRow + 1, 1)
    .getDisplayValues()
    .flat();

  const matrix = sheet
    .getRange(startRow, dateCol, lastRow - startRow + 1, numDateCols)
    .getDisplayValues();

  // Parse date headers into midnight dates
  const midnights = dateHeaders.map((hdr) => {
    const trimmed = String(hdr || "").trim();
    if (!trimmed) return null;

    // Try parseDate first (if headers look like "Jan 5, 2026")
    let d = Utilities.parseDate(trimmed, TIMEZONE, "MMM d, yyyy");
    if (isNaN(d)) {
      // Fallback: Date constructor
      d = new Date(trimmed);
    }
    if (isNaN(d)) return null;

    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  });

  const events = [];

  midnights.forEach((base, colIdx) => {
    if (!base) return;

    for (let rowIdx = 0; rowIdx < matrix.length; ) {
      const raw = String(matrix[rowIdx][colIdx] || "").trim();
      if (!raw) {
        rowIdx++;
        continue;
      }

      const timeStr = String(timeValues[rowIdx] || "").trim();
      const tm = timeStr.match(/(\d{1,2}):(\d{2})/);

      if (!tm) {
        Logger.log(`⚠️ No time match for rowOffset=${rowIdx}: "${timeStr}"`);
        rowIdx++;
        continue;
      }

      let [h, m] = tm.slice(1).map(Number);
      const originalH = h;

      // Your original PM-fix heuristic (keep)
      if (h >= 1 && h < 8) h += 12;

      const span =
        getMergedRowSpan(sheet, startRow + rowIdx, dateCol + colIdx) || 1;

      const startDT = new Date(base);
      startDT.setHours(h, m, 0, 0);

      const endDT = new Date(startDT);
      endDT.setMinutes(endDT.getMinutes() + span * 30);

      const title = normalizeTitle(raw);
      const isProject =
        PROJECT_CALL_TITLES.includes(raw) || raw.includes(UMBRAGE_KEYWORD);

      const key = buildEventKey(title, startDT, endDT);

      events.push({
        rawTitle: raw,
        title,
        start: startDT,
        end: endDT,
        span,
        timeStr,
        originalHour: originalH,
        isProject,
        key,
      });

      rowIdx += span;
    }
  });

  const projectKeys = new Set();
  const interviewKeys = new Set();
  events.forEach((evt) => (evt.isProject ? projectKeys : interviewKeys).add(evt.key));

  Logger.log(`🔑 Sheet keys: ${projectKeys.size} project, ${interviewKeys.size} interview`);
  return { events, projectKeys, interviewKeys };
}

function syncSheetEvents(events, projectCallCal, interviewsCal) {
  let createdCount = 0;
  let skippedCount = 0;

  events.forEach((evt) => {
    const cal = evt.isProject ? projectCallCal : interviewsCal;

    if (!cal) {
      Logger.log(`⚠️ Skipped (calendar missing): "${evt.title}" @ ${evt.start}`);
      return;
    }

    const exists = cal.getEvents(evt.start, evt.end).some((ev) => {
      return (
        normalizeTitle(ev.getTitle()) === evt.title &&
        ev.getStartTime().getTime() === evt.start.getTime() &&
        ev.getEndTime().getTime() === evt.end.getTime()
      );
    });

    const timeStrFormatted = Utilities.formatDate(evt.start, TIMEZONE, "MMM d, h:mm a");
    if (exists) {
      skippedCount++;
      Logger.log(`⏩ SKIPPED [${cal.getName()}] "${evt.title}" - ${timeStrFormatted} (already exists)`);
      return;
    }

    cal.createEvent(evt.title, evt.start, evt.end, {
      description: SCRIPT_TAG,
    });

    createdCount++;
    Logger.log(`✅ CREATED [${cal.getName()}] "${evt.title}" - ${timeStrFormatted}`);
  });

  return { createdCount, skippedCount };
}

function deleteRemovedEvents(calendar, sheetKeys, name) {
  if (!calendar) {
    Logger.log(`⚠️ Cannot clean stale for [${name}] - calendar not found`);
    return 0;
  }

  let deletedCount = 0;
  Logger.log(`🧹 Checking for stale events in [${name}]`);

  const now = new Date();
  const future = new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());

  calendar.getEvents(now, future).forEach((ev) => {
    const key = buildEventKey(normalizeTitle(ev.getTitle()), ev.getStartTime(), ev.getEndTime());
    if (!sheetKeys.has(key)) {
      const timeStr = Utilities.formatDate(ev.getStartTime(), TIMEZONE, "MMM d, h:mm a");
      ev.deleteEvent();
      deletedCount++;
      Logger.log(`🗑️ DELETED [${name}] "${ev.getTitle()}" - ${timeStr} (not in sheet)`);
    }
  });

  if (deletedCount === 0) Logger.log(`✓ No stale events to delete in [${name}]`);
  return deletedCount;
}

function getMergedRowSpan(sheet, row, col) {
  const r = sheet.getRange(row, col);
  if (!r.isPartOfMerge()) return 1;

  const merged = r
    .getMergedRanges()
    .find((m) => m.getRow() <= row && row <= m.getLastRow());

  return merged ? merged.getNumRows() : 1;
}
