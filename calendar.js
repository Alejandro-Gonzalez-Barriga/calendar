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
  dateColumn: 2, // column B
  dateColumnCount: 35,
  timeColumn: 1, // column A
});

// Safely normalize titles. If Apps Script runs this helper directly, run the sync.
function normalizeTitle(title) {
  if (typeof title === "undefined") {
    createCalendarEvents();
    return "";
  }

  const normalizedTitle = String(title || "").trim();
  return normalizedTitle.startsWith(UMBRAGE_KEYWORD) ? UMBRAGE_KEYWORD : normalizedTitle;
}

const buildEventKey = (title, startTime, endTime) =>
  `${title}|${startTime.getTime()}|${endTime.getTime()}`;

function log(message) {
  console.log(message);
}

function parseSheetTime(timeLabel, rawTimeValue) {
  if (rawTimeValue instanceof Date && !isNaN(rawTimeValue)) {
    const hour = rawTimeValue.getHours();
    const minute = rawTimeValue.getMinutes();
    return { hour, minute, originalHour: hour % 12 || 12, period: hour < 12 ? "AM" : "PM" };
  }

  if (typeof rawTimeValue === "number" && !isNaN(rawTimeValue)) {
    const minutesInDay = 24 * 60;
    const totalMinutes = Math.round((rawTimeValue % 1) * minutesInDay);
    const hour = Math.floor(totalMinutes / 60) % 24;
    const minute = totalMinutes % 60;
    return { hour, minute, originalHour: hour % 12 || 12, period: hour < 12 ? "AM" : "PM" };
  }

  const startTimeLabel = String(timeLabel || "").split("/")[0].trim();
  const timeMatch = startTimeLabel
    .trim()
    .match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]\.?m\.?)?$/i);

  if (!timeMatch) return null;

  let hour = Number(timeMatch[1]);
  const minute = Number(timeMatch[2] || 0);
  const period = timeMatch[3] ? timeMatch[3].replace(/\./g, "").toUpperCase() : "";
  const originalHour = hour;

  if (period === "AM") {
    if (hour === 12) hour = 0;
  } else if (period === "PM") {
    if (hour !== 12) hour += 12;
  } else if (hour >= 1 && hour < 7) {
    hour += 12;
  }

  if (hour > 23 || minute > 59) return null;

  return { hour, minute, originalHour, period };
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

function buildEventDescription(eventData) {
  const lines = [SCRIPT_TAG];

  if (eventData.note) {
    lines.push("", "Notes:", eventData.note);
  }

  if (eventData.links.length > 0) {
    lines.push("", "Links:");
    eventData.links.forEach((link) => {
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

    const { projectCallCalendar, interviewsCalendar } = getTargetCalendars();
    const { events, projectKeys, interviewKeys, parseFailureCount, syncStart, syncEnd } =
      extractEventsFromSheets(sheets, today);

    let deletedCount = 0;
    if (events.length === 0 || parseFailureCount > 0) {
      log(
        `⚠️ Skipping stale-event deletion: extracted=${events.length}, parseFailures=${parseFailureCount}`
      );
    } else {
      deletedCount =
        deleteRemovedEvents(projectCallCalendar, projectKeys, "Project call", syncStart, syncEnd) +
        deleteRemovedEvents(interviewsCalendar, interviewKeys, "Interviews", syncStart, syncEnd);
    }
    log(`📋 Total deleted: ${deletedCount}`);

    const { createdCount, skippedCount, updatedCount } = syncSheetEvents(
      events,
      projectCallCalendar,
      interviewsCalendar
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
  const spreadsheet = SpreadsheetApp.openById(SHEET_CONFIG.spreadsheetId);
  log("✅ Spreadsheet opened");

  const monthSheets = spreadsheet
    .getSheets()
    .map((sheet) => ({ sheet, monthStart: parseSheetMonth(sheet.getName()) }))
    .filter(({ monthStart }) => monthStart);

  const sheets = monthSheets
    .filter(({ monthStart }) => getMonthEndExclusive(monthStart) > today)
    .sort(
      (firstMonthSheet, secondMonthSheet) =>
        firstMonthSheet.monthStart.getTime() - secondMonthSheet.monthStart.getTime()
    )
    .map(({ sheet }) => sheet);

  if (sheets.length === 0) {
    log("⚠️ No current or future month sheets found. Aborting.");
    return [];
  }

  log(`✅ Current/future sheets found: ${sheets.map((sheet) => sheet.getName()).join(", ")}`);
  return sheets;
}

function getTargetCalendars() {
  const calendars = CalendarApp.getAllCalendars();
  const projectCallCalendar = calendars.find((calendar) => calendar.getName() === "Project call");
  const interviewsCalendar = calendars.find((calendar) => calendar.getName() === "Interviews");

  log(
    `📅 Using calendars: Project call='${projectCallCalendar ? projectCallCalendar.getName() : "NOT FOUND"}', ` +
      `Interviews='${interviewsCalendar ? interviewsCalendar.getName() : "NOT FOUND"}'`
  );

  return { projectCallCalendar, interviewsCalendar };
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
    result.projectKeys.forEach((eventKey) => projectKeys.add(eventKey));
    result.interviewKeys.forEach((eventKey) => interviewKeys.add(eventKey));
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
  const { startRow, lastRow, dateRow, dateColumn, dateColumnCount, timeColumn } =
    SHEET_CONFIG;

  const dateHeaders = sheet
    .getRange(dateRow, dateColumn, 1, dateColumnCount)
    .getDisplayValues()[0];

  const timeRange = sheet.getRange(startRow, timeColumn, lastRow - startRow + 1, 1);
  const timeDisplayValues = timeRange.getDisplayValues().flat();
  const timeRawValues = timeRange.getValues().flat();

  const eventRange = sheet.getRange(
    startRow,
    dateColumn,
    lastRow - startRow + 1,
    dateColumnCount
  );
  const eventDisplayValues = eventRange.getDisplayValues();
  const noteValues = eventRange.getNotes();
  const richTextValues = eventRange.getRichTextValues();
  const mergedRowSpans = getMergedRowSpans(eventRange, startRow, dateColumn);

  // Parse date headers into midnight dates
  const dateMidnights = dateHeaders.map((dateHeader) => {
    const trimmedDateHeader = String(dateHeader || "").trim();
    if (!trimmedDateHeader) return null;

    // Try parseDate first (if headers look like "Mar 5, 2026")
    let parsedDate = Utilities.parseDate(trimmedDateHeader, TIMEZONE, "MMM d, yyyy");
    if (isNaN(parsedDate)) {
      // Fallback: Date constructor
      parsedDate = new Date(trimmedDateHeader);
    }
    if (isNaN(parsedDate)) return null;

    return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
  });

  const events = [];
  let parseFailureCount = 0;
  let latestSheetDate = null;

  dateMidnights.forEach((eventDate, columnIndex) => {
    if (!eventDate) return;
    if (!latestSheetDate || eventDate > latestSheetDate) latestSheetDate = eventDate;
    if (eventDate < today) return;

    for (let rowIndex = 0; rowIndex < eventDisplayValues.length; ) {
      const rawTitle = String(eventDisplayValues[rowIndex][columnIndex] || "").trim();
      if (!rawTitle) {
        rowIndex++;
        continue;
      }

      const timeLabel = String(timeDisplayValues[rowIndex] || "").trim();
      const parsedTime = parseSheetTime(timeLabel, timeRawValues[rowIndex]);

      if (!parsedTime) {
        parseFailureCount++;
        log(`⚠️ No time match for rowOffset=${rowIndex}: "${timeLabel}"`);
        rowIndex++;
        continue;
      }

      const { hour, minute, originalHour } = parsedTime;

      const rowSpan = mergedRowSpans[`${rowIndex}:${columnIndex}`] || 1;

      const startDateTime = new Date(eventDate);
      startDateTime.setHours(hour, minute, 0, 0);

      const endDateTime = new Date(startDateTime);
      endDateTime.setMinutes(endDateTime.getMinutes() + rowSpan * 30);

      const title = normalizeTitle(rawTitle);
      const isProject =
        PROJECT_CALL_TITLES.includes(rawTitle) || rawTitle.includes(UMBRAGE_KEYWORD);

      const eventKey = buildEventKey(title, startDateTime, endDateTime);

      events.push({
        rawTitle,
        title,
        start: startDateTime,
        end: endDateTime,
        span: rowSpan,
        timeLabel,
        originalHour,
        note: String(noteValues[rowIndex][columnIndex] || "").trim(),
        links: getRichTextLinks(richTextValues[rowIndex][columnIndex]),
        isProject,
        key: eventKey,
      });

      rowIndex += rowSpan;
    }
  });

  const projectKeys = new Set();
  const interviewKeys = new Set();
  events.forEach((eventData) =>
    (eventData.isProject ? projectKeys : interviewKeys).add(eventData.key)
  );

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

function syncSheetEvents(events, projectCallCalendar, interviewsCalendar) {
  let createdCount = 0;
  let skippedCount = 0;
  let updatedCount = 0;

  events.forEach((eventData) => {
    const targetCalendar = eventData.isProject ? projectCallCalendar : interviewsCalendar;

    if (!targetCalendar) {
      log(`⚠️ Skipped (calendar missing): "${eventData.title}" @ ${eventData.start}`);
      return;
    }

    const existingEvent = targetCalendar.getEvents(eventData.start, eventData.end).find((calendarEvent) => {
      return (
        normalizeTitle(calendarEvent.getTitle()) === eventData.title &&
        calendarEvent.getStartTime().getTime() === eventData.start.getTime() &&
        calendarEvent.getEndTime().getTime() === eventData.end.getTime()
      );
    });

    const formattedStartTime = Utilities.formatDate(eventData.start, TIMEZONE, "MMM d, h:mm a");
    const description = buildEventDescription(eventData);
    if (existingEvent) {
      if (existingEvent.getDescription() !== description) {
        existingEvent.setDescription(description);
        updatedCount++;
        log(`📝 UPDATED [${targetCalendar.getName()}] "${eventData.title}" - ${formattedStartTime}`);
        return;
      }

      skippedCount++;
      log(`⏩ SKIPPED [${targetCalendar.getName()}] "${eventData.title}" - ${formattedStartTime} (already exists)`);
      return;
    }

    targetCalendar.createEvent(eventData.title, eventData.start, eventData.end, {
      description,
    });

    createdCount++;
    log(`✅ CREATED [${targetCalendar.getName()}] "${eventData.title}" - ${formattedStartTime}`);
  });

  return { createdCount, skippedCount, updatedCount };
}

function deleteRemovedEvents(calendar, expectedEventKeys, calendarName, syncStart, syncEnd) {
  if (!calendar) {
    log(`⚠️ Cannot clean stale for [${calendarName}] - calendar not found`);
    return 0;
  }

  let deletedCount = 0;
  const formattedSyncStart = Utilities.formatDate(syncStart, TIMEZONE, "MMM d, yyyy");
  const formattedSyncEnd = Utilities.formatDate(syncEnd, TIMEZONE, "MMM d, yyyy");
  log(`🧹 Checking for stale events in [${calendarName}] from ${formattedSyncStart} to ${formattedSyncEnd}`);

  calendar.getEvents(syncStart, syncEnd).forEach((calendarEvent) => {
    if (!String(calendarEvent.getDescription() || "").includes(SCRIPT_TAG)) return;

    const eventKey = buildEventKey(
      normalizeTitle(calendarEvent.getTitle()),
      calendarEvent.getStartTime(),
      calendarEvent.getEndTime()
    );
    if (!expectedEventKeys.has(eventKey)) {
      const formattedStartTime = Utilities.formatDate(calendarEvent.getStartTime(), TIMEZONE, "MMM d, h:mm a");
      calendarEvent.deleteEvent();
      deletedCount++;
      log(`🗑️ DELETED [${calendarName}] "${calendarEvent.getTitle()}" - ${formattedStartTime} (not in sheet)`);
    }
  });

  if (deletedCount === 0) log(`✓ No stale events to delete in [${calendarName}]`);
  return deletedCount;
}

function getMergedRowSpans(range, startRow, startColumn) {
  const spans = {};

  range.getMergedRanges().forEach((mergedRange) => {
    const rowOffset = mergedRange.getRow() - startRow;
    const columnOffset = mergedRange.getColumn() - startColumn;

    if (rowOffset < 0 || columnOffset < 0) return;
    spans[`${rowOffset}:${columnOffset}`] = mergedRange.getNumRows();
  });

  return spans;
}
