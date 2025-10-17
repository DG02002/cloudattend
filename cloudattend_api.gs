/**
 * CloudAttend Web App API
 * Handles RFID check-in/out flows and provides dashboard data.
 */

// Sheet configuration constants.
const CLOUDATTEND_DB_ID = "19jgGjDbxysLfbNrPufEvPOo9BJFK8-grxzY7VpW01WY";
const STUDENTS_SHEET_NAME = "Students";
const ATTENDANCE_SHEET_NAME = "Attendance";
const UNREGISTERED_SHEET_NAME = "Unregistered_CARDs";
const TIME_ZONE = "Asia/Kolkata";
const CORS_ALLOW_ORIGIN = "*";
const CORS_ALLOW_METHODS = "GET,POST,OPTIONS";
const CORS_ALLOW_HEADERS = "Content-Type";
const HUMAN_DATE_FORMAT = "d MMMM yyyy"; // e.g., 17 October 2025
const HUMAN_TIME_FORMAT = "h:mm a"; // e.g., 1:45 PM

/**
 * Handles POST requests from the RFID scanner.
 * @param {GoogleAppsScript.Events.DoPost} e
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Prevent concurrent writes from overlapping.

    const request = parseRequest(e);
    if (!request.valid) {
      return jsonResponse({ status: "error", message: request.message }, 400);
    }

    const action = (request.data.action || "").toString().toLowerCase();

    if (action === "scan") {
      const uid = (request.data.uid || "").toString().trim();
      if (!uid) {
        return jsonResponse(
          { status: "error", message: "UID is required" },
          400
        );
      }

      const now = new Date();
      const timeZone = TIME_ZONE;
      const isoTimestamp = Utilities.formatDate(
        now,
        timeZone,
        "yyyy-MM-dd'T'HH:mm:ssXXX"
      );
      const dateKey = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
      const humanDate = Utilities.formatDate(now, timeZone, HUMAN_DATE_FORMAT);
      const humanTime = Utilities.formatDate(now, timeZone, HUMAN_TIME_FORMAT);

      const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
      const attendanceSheet = getSheet(ATTENDANCE_SHEET_NAME);
      const unregisteredSheet = getSheet(UNREGISTERED_SHEET_NAME);

      const student = findStudentByUid(studentsSheet, uid);
      if (!student) {
        appendUnregistered(unregisteredSheet, uid, humanDate, humanTime);
        return jsonResponse({
          status: "ok",
          action: "unregistered",
          timestamp: isoTimestamp,
        });
      }

      const fullName = `${student.FirstName} ${student.LastName}`.trim();
      // Ensure headers exist and include a machine-usable DateKey plus human-readable Date/Time
      const attendanceHeaders = ensureAttendanceHeaders(attendanceSheet);
      const headerMap = attendanceHeaders.reduce((acc, header, idx) => {
        acc[header] = idx + 1; // 1-based for Range ops
        return acc;
      }, {});

      // Find latest open attendance row (today, same UID, without checkout)
      const openRow = findOpenAttendanceRowForDate(
        attendanceSheet,
        uid,
        dateKey
      );

      if (!openRow) {
        // No open session for today → create new check-in row
        const rowValues = [];
        rowValues[headerMap.CARD_UID - 1] = uid;
        rowValues[headerMap.SUID - 1] = student.SUID;
        rowValues[headerMap.Date - 1] = humanDate;
        rowValues[headerMap.DateKey - 1] = dateKey;
        rowValues[headerMap.CheckInTime - 1] = humanTime;
        rowValues[headerMap.CheckOutTime - 1] = "";
        rowValues[headerMap.Name - 1] = fullName;
        attendanceSheet.appendRow(rowValues);
        return jsonResponse({
          status: "ok",
          action: "checkin",
          timestamp: isoTimestamp,
          firstName: student.FirstName || "",
          fullName,
        });
      }

      // Open session exists → set checkout time
      const checkoutCol = openRow.headers.CheckOutTime;
      if (!checkoutCol) {
        throw new Error(
          "Attendance sheet headers must include CheckOutTime column"
        );
      }
      attendanceSheet.getRange(openRow.row, checkoutCol).setValue(humanTime);
      return jsonResponse({
        status: "ok",
        action: "checkout",
        timestamp: isoTimestamp,
        firstName: student.FirstName || "",
        fullName,
      });
    }

    if (action === "register") {
      const registrationPayload =
        request.data.data && typeof request.data.data === "object"
          ? request.data.data
          : request.data;
      const registrationResult = handleRegistration(registrationPayload);
      if (!registrationResult.success) {
        return jsonResponse(
          {
            status: "error",
            message: registrationResult.message,
          },
          registrationResult.statusCode || 400
        );
      }

      return jsonResponse(
        {
          status: "ok",
          action: "register",
          student: registrationResult.student,
        },
        200
      );
    }

    if (action === "delete") {
      const type = (
        (request.data.type || request.data.kind || "").toString() || ""
      ).toLowerCase();
      if (type === "unregistered") {
        const uid = (request.data.uid || request.data.cardUid || "")
          .toString()
          .trim()
          .toUpperCase();
        if (!uid) {
          return jsonResponse(
            { status: "error", message: "uid is required" },
            400
          );
        }
        const sheet = getSheet(UNREGISTERED_SHEET_NAME);
        const deleted = deleteUnregisteredByUid(sheet, uid);
        return jsonResponse({ status: "ok", action: "delete", deleted }, 200);
      }
      if (type === "attendance") {
        const uid = (request.data.uid || request.data.cardUid || "")
          .toString()
          .trim()
          .toUpperCase();
        const dateKey = (request.data.dateKey || "").toString().trim();
        const checkInTime = (
          request.data.checkInTime ||
          request.data.checkin ||
          ""
        )
          .toString()
          .trim();
        if (!uid || !dateKey || !checkInTime) {
          return jsonResponse(
            {
              status: "error",
              message: "uid, dateKey and checkInTime are required",
            },
            400
          );
        }
        const sheet = getSheet(ATTENDANCE_SHEET_NAME);
        ensureAttendanceHeaders(sheet);
        const removed = deleteAttendanceRow(sheet, uid, dateKey, checkInTime);
        return jsonResponse(
          { status: removed ? "ok" : "error", action: "delete", removed },
          removed ? 200 : 404
        );
      }
      return jsonResponse(
        { status: "error", message: "Unsupported delete type" },
        400
      );
    }

    return jsonResponse(
      { status: "error", message: "Unsupported action" },
      400
    );
  } catch (error) {
    return jsonResponse({ status: "error", message: error.message }, 500);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles GET requests for health checks, roster exports, or dashboard data.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.health === "1") {
    return handleHealthCheck();
  }

  if (params.registry === "1") {
    return handleRosterExport();
  }

  const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
  const attendanceSheet = getSheet(ATTENDANCE_SHEET_NAME);
  const unregisteredSheet = getSheet(UNREGISTERED_SHEET_NAME);

  // Ensure attendance headers and DateKey column exist so clients can filter today reliably
  ensureAttendanceHeaders(attendanceSheet);

  const response = {
    status: "ok",
    data: {
      students: sheetToObjects(studentsSheet),
      attendance: sheetToObjects(attendanceSheet),
      unregisteredCards: sheetToObjects(unregisteredSheet),
    },
  };

  return jsonResponse(response, 200);
}

/**
 * Handles CORS preflight requests.
 * @param {GoogleAppsScript.Events.DoPost} _e
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doOptions(_e) {
  return applyCorsHeaders(ContentService.createTextOutput(""));
}

/**
 * Lightweight endpoint for device health verification.
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function handleHealthCheck() {
  try {
    getSheet(STUDENTS_SHEET_NAME);
    getSheet(ATTENDANCE_SHEET_NAME);
    getSheet(UNREGISTERED_SHEET_NAME);
    return jsonResponse({ status: "ok" }, 200);
  } catch (error) {
    return jsonResponse({ status: "error", message: error.message }, 500);
  }
}

/**
 * Streams the student roster as a CSV payload.
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function handleRosterExport() {
  try {
    const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
    const csv = buildRosterCsv(studentsSheet);
    const output = ContentService.createTextOutput(csv).setMimeType(
      ContentService.MimeType.CSV
    );
    return applyCorsHeaders(output);
  } catch (error) {
    return jsonResponse({ status: "error", message: error.message }, 500);
  }
}

/**
 * Converts the Students sheet into a minimal CSV feed.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {string}
 */
function buildRosterCsv(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerLine = "CARD_UID,FirstName,LastName";
  if (!values.length) {
    return `${headerLine}\n`;
  }

  const headers = values.shift();
  const uidIndex = headers.indexOf("CARD_UID");
  const firstIndex = headers.indexOf("FirstName");
  const lastIndex = headers.indexOf("LastName");

  if (uidIndex === -1 || firstIndex === -1 || lastIndex === -1) {
    throw new Error(
      "Students sheet missing CARD_UID/FirstName/LastName headers"
    );
  }

  const lines = [headerLine];
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const rawUid = (row[uidIndex] || "").toString().trim();
    const rawFirst = (row[firstIndex] || "").toString().trim();
    const rawLast = (row[lastIndex] || "").toString().trim();

    if (!rawUid || !rawFirst) {
      continue;
    }

    const uid = rawUid.toUpperCase();
    lines.push(
      `${escapeCsv(uid)},${escapeCsv(rawFirst)},${escapeCsv(rawLast)}`
    );
  }

  return `${lines.join("\n")}\n`;
}

/**
 * Escapes CSV fields when needed.
 * @param {string} value
 * @return {string}
 */
function escapeCsv(value) {
  const str = (value || "").toString();
  if (/[",\n]/.test(str)) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * Parses the incoming POST payload into an object.
 * @param {GoogleAppsScript.Events.DoPost} e
 * @return {{ valid: boolean, data?: Object, message?: string }}
 */
function parseRequest(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return { valid: false, message: "Missing request body" };
  }

  try {
    const data = JSON.parse(e.postData.contents);
    if (!data || typeof data !== "object") {
      return { valid: false, message: "Invalid JSON payload" };
    }
    if (!data.action) {
      return { valid: false, message: "Action is required" };
    }
    return { valid: true, data };
  } catch (_err) {
    return { valid: false, message: "Invalid JSON payload" };
  }
}

/**
 * Retrieves the sheet by name within the CloudAttend_DB spreadsheet.
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(name) {
  const spreadsheet = SpreadsheetApp.openById(CLOUDATTEND_DB_ID);
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet not found: ${name}`);
  }
  return sheet;
}

/**
 * Finds a student record by CARD_UID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @return {Object|null}
 */
function findStudentByUid(sheet, uid) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return null; // No data rows available.
  }

  const headers = values.shift();
  const uidIndex = headers.indexOf("CARD_UID");
  if (uidIndex === -1) {
    throw new Error("CARD_UID column missing in Students sheet");
  }

  for (let i = 0; i < values.length; i += 1) {
    if (values[i][uidIndex] && values[i][uidIndex].toString() === uid) {
      return headers.reduce((acc, header, idx) => {
        acc[header] = values[i][idx];
        return acc;
      }, {});
    }
  }

  return null;
}

/**
 * Finds the first attendance row for a UID on the given date.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @param {string} dateKey
 * @return {{ row: number, headers: Object }|null}
 */
function findAttendanceRow(sheet, uid, dateKey) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return null;
  }

  const headers = values.shift();
  const headerMap = headers.reduce((acc, header, idx) => {
    acc[header] = idx + 1; // Store 1-based column indices for later use.
    return acc;
  }, {});

  const uidCol = headerMap.CARD_UID;
  const dateCol = headerMap.Date;

  if (!uidCol || !dateCol) {
    throw new Error("Attendance sheet headers must include CARD_UID and Date");
  }

  for (let i = 0; i < values.length; i += 1) {
    const rowValues = values[i];
    const rowUid = rowValues[uidCol - 1];
    const rowDate = rowValues[dateCol - 1];

    if (rowUid && rowUid.toString() === uid && rowDate === dateKey) {
      return { row: i + 2, headers: headerMap }; // +2 accounts for header row and 1-based indexing.
    }
  }

  return null;
}

/**
 * Finds the latest open (no checkout) attendance row for a UID on the dateKey.
 * Falls back to matching by Date column if legacy rows lack DateKey.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @param {string} dateKey
 * @return {{ row: number, headers: Object }|null}
 */
function findOpenAttendanceRowForDate(sheet, uid, dateKey) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return null;
  }

  const headers = values.shift();
  const headerMap = headers.reduce((acc, header, idx) => {
    acc[header] = idx + 1; // 1-based
    return acc;
  }, {});

  const uidCol = headerMap.CARD_UID;
  const dateKeyCol = headerMap.DateKey;
  const dateCol = headerMap.Date;
  const checkoutCol = headerMap.CheckOutTime;

  if (!uidCol || !dateCol || !checkoutCol) {
    throw new Error(
      "Attendance sheet headers must include CARD_UID, Date, and CheckOutTime"
    );
  }

  // Iterate from bottom (latest) upwards
  for (let i = values.length - 1; i >= 0; i -= 1) {
    const rowValues = values[i];
    const rowUid = (rowValues[uidCol - 1] || "").toString().trim();
    if (!rowUid || rowUid !== uid) {
      continue;
    }

    const primaryKey = dateKeyCol
      ? normalizeDateKeyValue(rowValues[dateKeyCol - 1])
      : "";
    let matchesDate = false;
    if (primaryKey) {
      matchesDate = primaryKey === dateKey;
    }

    if (!matchesDate) {
      const fallbackKey = normalizeDateKeyValue(rowValues[dateCol - 1]);
      matchesDate = fallbackKey === dateKey;
    }

    if (!matchesDate) {
      continue;
    }

    const checkoutValue = (rowValues[checkoutCol - 1] || "").toString().trim();
    if (!checkoutValue) {
      return { row: i + 2, headers: headerMap };
    }
    // If the latest matching row already has checkout, treat as no open row
    // and allow a new check-in by returning null.
    return null;
  }
  return null;
}

/**
 * Normalizes spreadsheet date-like values into ISO yyyy-MM-dd strings.
 * @param {any} value
 * @return {string}
 */
function normalizeDateKeyValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return Utilities.formatDate(value, TIME_ZONE, "yyyy-MM-dd");
  }

  const str = value.toString().trim();
  if (!str) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, TIME_ZONE, "yyyy-MM-dd");
  }

  return "";
}

/**
 * Appends an entry to the Unregistered_CARDs sheet, adding a Status column if present.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @param {string} dateKey
 * @param {string} timeKey
 */
function appendUnregistered(sheet, uid, dateKey, timeKey) {
  const statusColumn = ensureStatusColumn(sheet);
  const row = [uid, dateKey, timeKey];

  if (statusColumn) {
    row.push("Pending");
  }

  sheet.appendRow(row);
}

/**
 * Handles registration requests coming from the dashboard.
 * @param {Object} data
 * @return {{ success: boolean, message?: string, statusCode?: number, student?: Object }}
 */
function handleRegistration(data) {
  if (!data || typeof data !== "object") {
    return {
      success: false,
      message: "Invalid registration payload",
      statusCode: 400,
    };
  }

  const cardUid = (data.cardUid || data.uid || "")
    .toString()
    .trim()
    .toUpperCase();
  const suid = (data.suid || "").toString().trim();
  const firstName = (data.firstName || "").toString().trim();
  const lastName = (data.lastName || "").toString().trim();

  const missing = [];
  if (!cardUid) missing.push("cardUid");
  if (!suid) missing.push("suid");
  if (!firstName) missing.push("firstName");
  if (!lastName) missing.push("lastName");

  if (missing.length) {
    return {
      success: false,
      message: `Missing fields: ${missing.join(", ")}`,
      statusCode: 400,
    };
  }

  const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
  const unregisteredSheet = getSheet(UNREGISTERED_SHEET_NAME);

  const normalizedStudent = {
    CARD_UID: cardUid,
    SUID: suid,
    FirstName: firstName,
    LastName: lastName,
  };

  const upsertResult = upsertStudent(studentsSheet, normalizedStudent);
  markCardAsRegistered(unregisteredSheet, cardUid);

  return {
    success: true,
    student: upsertResult.student,
  };
}

/**
 * Inserts or updates a student in the Students sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} record
 * @return {{ created: boolean, student: Object }}
 */
function upsertStudent(sheet, record) {
  const headers = ensureStudentHeaders(sheet);
  const headerMap = headers.reduce((acc, header, idx) => {
    if (header) {
      acc[header] = idx;
    }
    return acc;
  }, {});

  if (headerMap.CARD_UID === undefined) {
    throw new Error("Students sheet missing CARD_UID header");
  }

  const lastColumn = headers.length;
  const rowValues = headers.map((header) => record[header] || "");
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    const values = dataRange.getValues();
    for (let i = 0; i < values.length; i += 1) {
      const rowUid = values[i][headerMap.CARD_UID];
      if (
        rowUid &&
        rowUid.toString().trim().toUpperCase() === record.CARD_UID.toUpperCase()
      ) {
        sheet.getRange(i + 2, 1, 1, lastColumn).setValues([rowValues]);
        return { created: false, student: record };
      }
    }
  }

  sheet.appendRow(rowValues);
  return { created: true, student: record };
}

/**
 * Ensures the Students sheet has the expected headers and returns them.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array<string>}
 */
function ensureStudentHeaders(sheet) {
  const expectedHeaders = ["CARD_UID", "SUID", "FirstName", "LastName"];

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet
      .getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders]);
    return expectedHeaders;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  for (let i = 0; i < expectedHeaders.length; i += 1) {
    if (headers[i] !== expectedHeaders[i]) {
      sheet.getRange(1, i + 1).setValue(expectedHeaders[i]);
      headers[i] = expectedHeaders[i];
    }
  }

  return headers.slice(0, expectedHeaders.length);
}

/**
 * Ensures the Unregistered_CARDs sheet contains a Status column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number} 1-based column index of Status
 */
function ensureStatusColumn(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, 4)
      .setValues([["CARD_UID", "Date", "Time", "Status"]]);
    return 4;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let statusIndex = headers.indexOf("Status");
  if (statusIndex !== -1) {
    return statusIndex + 1;
  }

  const newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue("Status");
  return newColumn;
}

/**
 * Updates the Unregistered_CARDs sheet to mark a card as registered.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} cardUid
 * @return {boolean}
 */
function markCardAsRegistered(sheet, cardUid) {
  const statusColumn = ensureStatusColumn(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const uidIndex = headers.indexOf("CARD_UID");
  if (uidIndex === -1) {
    return false;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  for (let i = 0; i < values.length; i += 1) {
    const rowUid = values[i][uidIndex];
    if (rowUid && rowUid.toString().trim().toUpperCase() === cardUid) {
      sheet.getRange(i + 2, statusColumn).setValue("Registered");
      return true;
    }
  }

  return false;
}

/**
 * Converts a sheet into an array of objects keyed by header names.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array<Object>}
 */
function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const headers = values.shift();
  return values
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) =>
      headers.reduce((acc, header, idx) => {
        acc[header] = row[idx];
        return acc;
      }, {})
    );
}

/**
 * Builds a JSON HTTP response.
 * @param {Object} payload
 * @param {number} statusCode
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function jsonResponse(payload, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  // Apps Script Web Apps do not support setting HTTP status codes on TextOutput.
  // We encode status in the JSON payload instead.
  return applyCorsHeaders(output);
}

/**
 * Adds CORS headers to the outgoing response.
 * @param {GoogleAppsScript.Content.TextOutput} output
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function applyCorsHeaders(output) {
  // Note: Apps Script ContentService.TextOutput doesn't support setting headers.
  // For Web Apps, responses are same-origin to the web app URL, so CORS headers are not required.
  // If you need CORS for external origins, consider using HtmlService and returning a templated HTML that fetches internally.
  return output;
}

/**
 * Ensures the Attendance sheet has the expected headers and returns them.
 * Expected columns (order enforced for new sheets):
 * CARD_UID | SUID | Date | DateKey | CheckInTime | CheckOutTime | Name
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array<string>}
 */
function ensureAttendanceHeaders(sheet) {
  const expected = [
    "CARD_UID",
    "SUID",
    "Date",
    "DateKey",
    "CheckInTime",
    "CheckOutTime",
    "Name",
  ];
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return expected;
  }

  let lastColumn = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  // Ensure DateKey column exists; if missing, insert after Date and backfill
  let dateIndex = headers.indexOf("Date");
  let dateKeyIndex = headers.indexOf("DateKey");
  if (dateKeyIndex === -1) {
    if (dateIndex === -1) {
      // If Date column is missing, append both Date and DateKey at the end.
      sheet.insertColumnsAfter(lastColumn, 2);
      sheet.getRange(1, lastColumn + 1).setValue("Date");
      sheet.getRange(1, lastColumn + 2).setValue("DateKey");
      // No backfill possible without Date; leave blank.
      lastColumn += 2;
    } else {
      // Insert DateKey right after Date
      sheet.insertColumnAfter(dateIndex + 1);
      sheet.getRange(1, dateIndex + 2).setValue("DateKey");
      lastColumn = sheet.getLastColumn();
      // Backfill DateKey using Date column values
      const dataLastRow = sheet.getLastRow();
      if (dataLastRow > 1) {
        const dateRange = sheet.getRange(2, dateIndex + 1, dataLastRow - 1, 1);
        const dateValues = dateRange.getValues();
        const out = [];
        for (let i = 0; i < dateValues.length; i++) {
          const cell = dateValues[i][0];
          let key = "";
          if (cell instanceof Date) {
            key = Utilities.formatDate(cell, TIME_ZONE, "yyyy-MM-dd");
          } else {
            const str = (cell || "").toString();
            if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
              key = str;
            } else {
              const parsed = new Date(str);
              if (!isNaN(parsed.getTime())) {
                key = Utilities.formatDate(parsed, TIME_ZONE, "yyyy-MM-dd");
              }
            }
          }
          out.push([key]);
        }
        sheet.getRange(2, dateIndex + 2, out.length, 1).setValues(out);
      }
    }
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // Ensure other expected headers exist; if missing, append new columns with header
  const headerSet = new Set(headers.filter((h) => !!h));
  for (let i = 0; i < expected.length; i++) {
    const name = expected[i];
    if (!headerSet.has(name)) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      const col = sheet.getLastColumn();
      sheet.getRange(1, col).setValue(name);
      headerSet.add(name);
    }
  }

  const finalHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  return finalHeaders;
}

/**
 * Deletes unregistered rows by UID (case-insensitive). Returns count deleted.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @return {number}
 */
function deleteUnregisteredByUid(sheet, uid) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const uidIndex = headers.indexOf("CARD_UID");
  if (uidIndex === -1) return 0;
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = range.getValues();
  let count = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const v = (values[i][uidIndex] || "").toString().trim().toUpperCase();
    if (v === uid) {
      sheet.deleteRow(i + 2);
      count++;
    }
  }
  return count;
}

/**
 * Deletes a single attendance row matching UID + DateKey + CheckInTime.
 * Returns true if a row was deleted.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} uid
 * @param {string} dateKey
 * @param {string} checkInTime
 */
function deleteAttendanceRow(sheet, uid, dateKey, checkInTime) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = headers.reduce((acc, h, i) => {
    acc[h] = i;
    return acc;
  }, {});
  const uidIdx = headerMap.CARD_UID;
  const dateKeyIdx = headerMap.DateKey;
  const dateIdx = headerMap.Date;
  const inIdx = headerMap.CheckInTime;
  if (uidIdx === undefined || inIdx === undefined) return false;

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = range.getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const rowUid = (row[uidIdx] || "").toString().trim().toUpperCase();
    if (rowUid !== uid) continue;
    const rowCheckIn = (row[inIdx] || "").toString().trim();
    // Match CheckInTime exactly (human readable or HH:mm)
    if (rowCheckIn !== checkInTime) continue;

    // Match DateKey, fallback to parsing Date column
    let rowDateKey =
      dateKeyIdx !== undefined ? (row[dateKeyIdx] || "").toString().trim() : "";
    if (!rowDateKey && dateIdx !== undefined) {
      const d = row[dateIdx];
      if (d instanceof Date) {
        rowDateKey = Utilities.formatDate(d, TIME_ZONE, "yyyy-MM-dd");
      } else {
        const s = (d || "").toString();
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
          rowDateKey = s;
        } else {
          const parsed = new Date(s);
          if (!isNaN(parsed.getTime())) {
            rowDateKey = Utilities.formatDate(parsed, TIME_ZONE, "yyyy-MM-dd");
          }
        }
      }
    }

    if (rowDateKey === dateKey) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}
