// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS & CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────────

/** Keyword used to normalize "Umbrage" prefixed titles. */
const UMBRAGE_KEYWORD = "Umbrage";
/** Titles that belong to the "Project call" calendar vs "Interviews". */
const PROJECT_CALL_TITLES = ["Project call", "Resla", "CheckSammy", "RevStar"];
/** Timezone used for parsing and formatting dates. */
const TIMEZONE = "America/Mexico_City";
/** Description tag added to events created by this script. */
const SCRIPT_TAG = "Created by Script";

/** Spreadsheet layout: IDs, sheet name, row/column ranges for dates and times. */
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

// ─────────────────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

/** Normalizes a title: trims whitespace and collapses Umbrage-prefixed titles to the keyword. */
const normalizeTitle = (title = "") => {
  const str = String(title || "").trim();
  return str.startsWith(UMBRAGE_KEYWORD) ? UMBRAGE_KEYWORD : str;
};

/** Builds a unique string key for an event (title|startMs|endMs). */
const buildEventKey = (title, startTime, endTime) =>
  `${title}|${startTime.getTime()}|${endTime.getTime()}`;

// ─────────────────────────────────────────────────────────────────────────────
// MAIN ORCHESTRATION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Main entry point. Syncs calendar events from the configured Google Sheet to
 * "Project call" and "Interviews" calendars: creates new events, skips existing
 * ones, and deletes events that no longer appear in the sheet.
 */
function createCalendarEvents() {
  try {
    Logger.log("🔄 Starting calendar sync");

    const sheet = getTargetSheet();
    if (!sheet) return;

    const { projectCallCal, interviewsCal } = getTargetCalendars();
    const { events, projectKeys, interviewKeys } = extractSheetEvents(sheet);

    // Delete events that exist in calendar but not in sheet
    const deletedCount =
      deleteRemovedEvents(projectCallCal, projectKeys, "Project call") +
      deleteRemovedEvents(interviewsCal, interviewKeys, "Interviews");
    Logger.log(`📋 Total deleted: ${deletedCount}`);

    // Create new events (skip duplicates)
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

/**
 * Opens the spreadsheet by ID and returns the sheet specified in SHEET_CONFIG.
 * Returns null if the sheet is not found.
 */
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

/**
 * Finds the "Project call" and "Interviews" calendars from the user's list.
 * Returns both calendar objects (may be undefined if not found).
 */
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

/**
 * Reads the sheet matrix and extracts event data: title, start/end times, span,
 * and whether each event is a project call or interview. Also returns sets of
 * keys for project and interview events (used for deletion of stale events).
 */
function extractSheetEvents(sheet) {
  const { startRow, lastRow, dateRow, dateCol, numDateCols, timeCol } =
    SHEET_CONFIG;

  // Read date headers (e.g., "Jan 5, 2026") from the header row
  const dateHeaders = sheet
    .getRange(dateRow, dateCol, 1, numDateCols)
    .getDisplayValues()[0];

  // Read time values from column A (start times for each row)
  const timeValues = sheet
    .getRange(startRow, timeCol, lastRow - startRow + 1, 1)
    .getDisplayValues()
    .flat();

  // Read the cell matrix (event titles per date/row)
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

  // Iterate each date column; for each non-empty cell, parse time and build event
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

  // Split events into project vs interview key sets for delete logic
  const projectKeys = new Set();
  const interviewKeys = new Set();
  events.forEach((evt) => (evt.isProject ? projectKeys : interviewKeys).add(evt.key));

  Logger.log(`🔑 Sheet keys: ${projectKeys.size} project, ${interviewKeys.size} interview`);
  return { events, projectKeys, interviewKeys };
}

/**
 * Creates calendar events from the extracted sheet events. Skips events that
 * already exist (same title, start, end). Returns counts of created and skipped.
 */
function syncSheetEvents(events, projectCallCal, interviewsCal) {
  let createdCount = 0;
  let skippedCount = 0;

  // For each event: skip if calendar missing or event already exists; otherwise create
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

/**
 * Deletes calendar events that are not in the given sheetKeys set. Used to
 * remove events that were previously created but have been removed from the
 * sheet. Returns the number of events deleted.
 */
function deleteRemovedEvents(calendar, sheetKeys, name) {
  if (!calendar) {
    Logger.log(`⚠️ Cannot clean stale for [${name}] - calendar not found`);
    return 0;
  }

  let deletedCount = 0;
  Logger.log(`🧹 Checking for stale events in [${name}]`);

  const now = new Date();
  const future = new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());

  // Delete any calendar events in the next month that are not in the sheet
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

/**
 * Returns the row span of a merged cell at (row, col). If the cell is part of
 * a merge, returns the number of rows in that merge; otherwise returns 1.
 */
function getMergedRowSpan(sheet, row, col) {
  const r = sheet.getRange(row, col);
  if (!r.isPartOfMerge()) return 1;

  // Find the merge range that contains this cell
  const merged = r
    .getMergedRanges()
    .find((m) => m.getRow() <= row && row <= m.getLastRow());

  return merged ? merged.getNumRows() : 1;
}
