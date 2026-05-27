// Constants
const UMBRAGE_KEYWORD = "Umbrage";
const PROJECT_CALL_TITLES = ["Project call", "Resla", "CheckSammy", "RevStar"];
const TIMEZONE = "America/Mexico_City";
const SCRIPT_TAG = "Created by Script";
const MONTH_SHEET_NAME_PATTERN = /^([A-Za-z]+)\s+(\d{4})$/;
const MONTH_NAME_TO_INDEX = Object.freeze({
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
});

const SHEET_CONFIG = Object.freeze({
  spreadsheetId: "1AMGUOTTL3cVrhNcy55xRIDFllFoBAgL5iRkjvWXAa50",
  startRow: 11,
  lastRow: 37,
  dateRow: 3,
  dateCol: 2,      // column B
  numDateCols: 35,
  timeCol: 1,      // column A
});

// Safely normalize titles. If Apps Script runs this helper directly, run the sync.
function normalizeTitle(title) {
  if (typeof title === "undefined") {
    createCalendarEvents();
    return "";
  }

  const str = String(title || "").trim();
  return str.startsWith(UMBRAGE_KEYWORD) ? UMBRAGE_KEYWORD : str;
}

const buildEventKey = (title, startTime, endTime) =>
  `${title}|${startTime.getTime()}|${endTime.getTime()}`;

function log(message) {
  console.log(message);
}

function parseSheetTime(timeStr, rawValue) {
  if (rawValue instanceof Date && !isNaN(rawValue)) {
    const h = rawValue.getHours();
    const m = rawValue.getMinutes();
    return { h, m, originalHour: h % 12 || 12, period: h < 12 ? "AM" : "PM" };
  }

  if (typeof rawValue === "number" && !isNaN(rawValue)) {
    const minutesInDay = 24 * 60;
    const totalMinutes = Math.round((rawValue % 1) * minutesInDay);
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = totalMinutes % 60;
    return { h, m, originalHour: h % 12 || 12, period: h < 12 ? "AM" : "PM" };
  }

  const startTimeStr = String(timeStr || "").split("/")[0].trim();
  const tm = startTimeStr
    .trim()
    .match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]\.?m\.?)?$/i);

  if (!tm) return null;

  let h = Number(tm[1]);
  const m = Number(tm[2] || 0);
  const period = tm[3] ? tm[3].replace(/\./g, "").toUpperCase() : "";
  const originalHour = h;

  if (period === "AM") {
    if (h === 12) h = 0;
  } else if (period === "PM") {
    if (h !== 12) h += 12;
  } else if (h >= 1 && h < 7) {
    h += 12;
  }

  if (h > 23 || m > 59) return null;

  return { h, m, originalHour, period };
}

function getRichTextLinks(richTextValue) {
  if (!richTextValue) return [];

  const links = [];
  const fullText = richTextValue.getText ? richTextValue.getText() : "";
  const fullLink = richTextValue.getLinkUrl ? richTextValue.getLinkUrl() : null;

  if (fullLink) {
    links.push({ text: fullText || fullLink, url: fullLink });
  }

  const runs = richTextValue.getRuns ? richTextValue.getRuns() : [];
  runs.forEach((run) => {
    const url = run.getLinkUrl ? run.getLinkUrl() : null;
    if (!url || links.some((link) => link.url === url)) return;

    const text = run.getText ? run.getText() : url;
    links.push({ text: text || url, url });
  });

  return links;
}

function buildEventDescription(evt) {
  const lines = [SCRIPT_TAG];

  if (evt.note) {
    lines.push("", "Notes:", evt.note);
  }

  if (evt.links.length > 0) {
    lines.push("", "Links:");
    evt.links.forEach((link) => {
      lines.push(`${link.text}: ${link.url}`);
    });
  }

  return lines.join("\n");
}

function myFunction() {
  createCalendarEvents();
}

function createCalendarEvents() {
  try {
    log("🔄 Starting calendar sync");

    const today = getTodayMidnight();
    const sheets = getTargetSheets(today);
    if (sheets.length === 0) return;

    const { projectCallCal, interviewsCal } = getTargetCalendars();
    const { events, projectKeys, interviewKeys, parseFailureCount, syncStart, syncEnd } =
      extractEventsFromSheets(sheets, today);

    let deletedCount = 0;
    if (events.length === 0 || parseFailureCount > 0) {
      log(
        `⚠️ Skipping stale-event deletion: extracted=${events.length}, parseFailures=${parseFailureCount}`
      );
    } else {
      deletedCount =
        deleteRemovedEvents(projectCallCal, projectKeys, "Project call", syncStart, syncEnd) +
        deleteRemovedEvents(interviewsCal, interviewKeys, "Interviews", syncStart, syncEnd);
    }
    log(`📋 Total deleted: ${deletedCount}`);

    const { createdCount, skippedCount, updatedCount } = syncSheetEvents(
      events,
      projectCallCal,
      interviewsCal
    );

    log("=".repeat(50));
    log(
      `📊 SUMMARY: Created=${createdCount}, Updated=${updatedCount}, Skipped=${skippedCount}, Deleted=${deletedCount}`
    );
    log("✅ Calendar sync completed");
  } catch (error) {
    log(`❌ ERROR: ${error.toString()}`);
    log(`Stack: ${error.stack}`);
    throw error;
  }
}

function getTodayMidnight() {
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  return Utilities.parseDate(todayStr, TIMEZONE, "yyyy-MM-dd");
}

function parseSheetMonth(sheetName) {
  const match = String(sheetName || "")
    .trim()
    .match(MONTH_SHEET_NAME_PATTERN);
  if (!match) return null;

  const monthToken = match[1].toLowerCase();

  if (!(monthToken in MONTH_NAME_TO_INDEX)) return null;
  return new Date(Number(match[2]), MONTH_NAME_TO_INDEX[monthToken], 1);
}

function getMonthEndExclusive(monthStart) {
  return new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 1);
}

function getTargetSheets(today) {
  const ss = SpreadsheetApp.openById(SHEET_CONFIG.spreadsheetId);
  log("✅ Spreadsheet opened");

  const monthSheets = ss
    .getSheets()
    .map((sheet) => ({ sheet, monthStart: parseSheetMonth(sheet.getName()) }))
    .filter(({ monthStart }) => monthStart);

  const sheets = monthSheets
    .filter(({ monthStart }) => getMonthEndExclusive(monthStart) > today)
    .sort((a, b) => a.monthStart.getTime() - b.monthStart.getTime())
    .map(({ sheet }) => sheet);

  if (sheets.length === 0) {
    log("⚠️ No current or future month sheets found. Aborting.");
    return [];
  }

  log(`✅ Current/future sheets found: ${sheets.map((sheet) => sheet.getName()).join(", ")}`);
  return sheets;
}

function getTargetCalendars() {
  const allCals = CalendarApp.getAllCalendars();
  const projectCallCal = allCals.find((c) => c.getName() === "Project call");
  const interviewsCal = allCals.find((c) => c.getName() === "Interviews");

  log(
    `📅 Using calendars: Project call='${projectCallCal ? projectCallCal.getName() : "NOT FOUND"}', ` +
      `Interviews='${interviewsCal ? interviewsCal.getName() : "NOT FOUND"}'`
  );

  return { projectCallCal, interviewsCal };
}

function extractEventsFromSheets(sheets, today) {
  const combinedEvents = [];
  const projectKeys = new Set();
  const interviewKeys = new Set();
  let parseFailureCount = 0;
  let syncStart = today;
  let syncEnd = null;

  sheets.forEach((sheet) => {
    const result = extractSheetEvents(sheet, today);
    combinedEvents.push(...result.events);
    result.projectKeys.forEach((key) => projectKeys.add(key));
    result.interviewKeys.forEach((key) => interviewKeys.add(key));
    parseFailureCount += result.parseFailureCount;

    if (result.syncEnd && (!syncEnd || result.syncEnd > syncEnd)) {
      syncEnd = result.syncEnd;
    }
  });

  if (!syncEnd) {
    syncEnd = new Date(today.getFullYear(), today.getMonth() + 1, today.getDate());
  }

  log(
    `🔑 Sheet keys: ${projectKeys.size} project, ${interviewKeys.size} interview, parse failures=${parseFailureCount}`
  );
  return {
    events: combinedEvents,
    projectKeys,
    interviewKeys,
    parseFailureCount,
    syncStart,
    syncEnd,
  };
}

function extractSheetEvents(sheet, today) {
  const { startRow, lastRow, dateRow, dateCol, numDateCols, timeCol } =
    SHEET_CONFIG;

  const dateHeaders = sheet
    .getRange(dateRow, dateCol, 1, numDateCols)
    .getDisplayValues()[0];

  const timeRange = sheet.getRange(startRow, timeCol, lastRow - startRow + 1, 1);
  const timeDisplayValues = timeRange.getDisplayValues().flat();
  const timeRawValues = timeRange.getValues().flat();

  const eventRange = sheet.getRange(
    startRow,
    dateCol,
    lastRow - startRow + 1,
    numDateCols
  );
  const matrix = eventRange.getDisplayValues();
  const noteMatrix = eventRange.getNotes();
  const richTextMatrix = eventRange.getRichTextValues();
  const mergedRowSpans = getMergedRowSpans(eventRange, startRow, dateCol);

  // Parse date headers into midnight dates
  const midnights = dateHeaders.map((hdr) => {
    const trimmed = String(hdr || "").trim();
    if (!trimmed) return null;

    // Try parseDate first (if headers look like "Mar 5, 2026")
    let d = Utilities.parseDate(trimmed, TIMEZONE, "MMM d, yyyy");
    if (isNaN(d)) {
      // Fallback: Date constructor
      d = new Date(trimmed);
    }
    if (isNaN(d)) return null;

    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  });

  const events = [];
  let parseFailureCount = 0;
  let latestSheetDate = null;

  midnights.forEach((base, colIdx) => {
    if (!base) return;
    if (!latestSheetDate || base > latestSheetDate) latestSheetDate = base;
    if (base < today) return;

    for (let rowIdx = 0; rowIdx < matrix.length; ) {
      const raw = String(matrix[rowIdx][colIdx] || "").trim();
      if (!raw) {
        rowIdx++;
        continue;
      }

      const timeStr = String(timeDisplayValues[rowIdx] || "").trim();
      const parsedTime = parseSheetTime(timeStr, timeRawValues[rowIdx]);

      if (!parsedTime) {
        parseFailureCount++;
        log(`⚠️ No time match for rowOffset=${rowIdx}: "${timeStr}"`);
        rowIdx++;
        continue;
      }

      const { h, m, originalHour } = parsedTime;

      const span = mergedRowSpans[`${rowIdx}:${colIdx}`] || 1;

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
        originalHour,
        note: String(noteMatrix[rowIdx][colIdx] || "").trim(),
        links: getRichTextLinks(richTextMatrix[rowIdx][colIdx]),
        isProject,
        key,
      });

      rowIdx += span;
    }
  });

  const projectKeys = new Set();
  const interviewKeys = new Set();
  events.forEach((evt) => (evt.isProject ? projectKeys : interviewKeys).add(evt.key));

  log(
    `📄 ${sheet.getName()}: ${events.length} events, ${projectKeys.size} project, ${interviewKeys.size} interview, parse failures=${parseFailureCount}`
  );
  return {
    events,
    projectKeys,
    interviewKeys,
    parseFailureCount,
    syncEnd: latestSheetDate
      ? new Date(latestSheetDate.getFullYear(), latestSheetDate.getMonth(), latestSheetDate.getDate() + 1)
      : null,
  };
}

function syncSheetEvents(events, projectCallCal, interviewsCal) {
  let createdCount = 0;
  let skippedCount = 0;
  let updatedCount = 0;

  events.forEach((evt) => {
    const cal = evt.isProject ? projectCallCal : interviewsCal;

    if (!cal) {
      log(`⚠️ Skipped (calendar missing): "${evt.title}" @ ${evt.start}`);
      return;
    }

    const existingEvent = cal.getEvents(evt.start, evt.end).find((ev) => {
      return (
        normalizeTitle(ev.getTitle()) === evt.title &&
        ev.getStartTime().getTime() === evt.start.getTime() &&
        ev.getEndTime().getTime() === evt.end.getTime()
      );
    });

    const timeStrFormatted = Utilities.formatDate(evt.start, TIMEZONE, "MMM d, h:mm a");
    const description = buildEventDescription(evt);
    if (existingEvent) {
      if (existingEvent.getDescription() !== description) {
        existingEvent.setDescription(description);
        updatedCount++;
        log(`📝 UPDATED [${cal.getName()}] "${evt.title}" - ${timeStrFormatted}`);
        return;
      }

      skippedCount++;
      log(`⏩ SKIPPED [${cal.getName()}] "${evt.title}" - ${timeStrFormatted} (already exists)`);
      return;
    }

    cal.createEvent(evt.title, evt.start, evt.end, {
      description,
    });

    createdCount++;
    log(`✅ CREATED [${cal.getName()}] "${evt.title}" - ${timeStrFormatted}`);
  });

  return { createdCount, skippedCount, updatedCount };
}

function deleteRemovedEvents(calendar, sheetKeys, name, syncStart, syncEnd) {
  if (!calendar) {
    log(`⚠️ Cannot clean stale for [${name}] - calendar not found`);
    return 0;
  }

  let deletedCount = 0;
  const startStr = Utilities.formatDate(syncStart, TIMEZONE, "MMM d, yyyy");
  const endStr = Utilities.formatDate(syncEnd, TIMEZONE, "MMM d, yyyy");
  log(`🧹 Checking for stale events in [${name}] from ${startStr} to ${endStr}`);

  calendar.getEvents(syncStart, syncEnd).forEach((ev) => {
    if (!String(ev.getDescription() || "").includes(SCRIPT_TAG)) return;

    const key = buildEventKey(normalizeTitle(ev.getTitle()), ev.getStartTime(), ev.getEndTime());
    if (!sheetKeys.has(key)) {
      const timeStr = Utilities.formatDate(ev.getStartTime(), TIMEZONE, "MMM d, h:mm a");
      ev.deleteEvent();
      deletedCount++;
      log(`🗑️ DELETED [${name}] "${ev.getTitle()}" - ${timeStr} (not in sheet)`);
    }
  });

  if (deletedCount === 0) log(`✓ No stale events to delete in [${name}]`);
  return deletedCount;
}

function getMergedRowSpans(range, startRow, startCol) {
  const spans = {};

  range.getMergedRanges().forEach((mergedRange) => {
    const rowOffset = mergedRange.getRow() - startRow;
    const colOffset = mergedRange.getColumn() - startCol;

    if (rowOffset < 0 || colOffset < 0) return;
    spans[`${rowOffset}:${colOffset}`] = mergedRange.getNumRows();
  });

  return spans;
}
