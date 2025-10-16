/**
 * CloudAttend Web App API
 * Handles RFID check-in/out flows and provides dashboard data.
 */

// Sheet configuration constants.
const CLOUDATTEND_DB = "CloudAttend_DB";
const STUDENTS_SHEET_NAME = "Students";
const ATTENDANCE_SHEET_NAME = "Attendance";
const UNREGISTERED_SHEET_NAME = "Unregistered_CARDs";

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
      const timeZone = Session.getScriptTimeZone();
      const isoTimestamp = Utilities.formatDate(
        now,
        timeZone,
        "yyyy-MM-dd'T'HH:mm:ssXXX"
      );
      const currentDate = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
      const currentTime = Utilities.formatDate(now, timeZone, "HH:mm:ss");

      const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
      const attendanceSheet = getSheet(ATTENDANCE_SHEET_NAME);
      const unregisteredSheet = getSheet(UNREGISTERED_SHEET_NAME);

      const student = findStudentByUid(studentsSheet, uid);
      if (!student) {
        appendUnregistered(unregisteredSheet, uid, currentDate, currentTime);
        return jsonResponse({
          status: "ok",
          action: "unregistered",
          timestamp: isoTimestamp,
        });
      }

      const fullName = `${student.FirstName} ${student.LastName}`.trim();
      const existingRow = findAttendanceRow(attendanceSheet, uid, currentDate);

      if (!existingRow) {
        // First scan of the day for this UID → create check-in row.
        attendanceSheet.appendRow([
          uid,
          student.SUID,
          currentDate,
          currentTime,
          "",
          fullName,
        ]);
        return jsonResponse({
          status: "ok",
          action: "checkin",
          timestamp: isoTimestamp,
        });
      }

      const checkoutCell = attendanceSheet.getRange(
        existingRow.row,
        existingRow.headers.CheckOutTime
      );
      const checkoutValue = checkoutCell.getValue();

      if (!checkoutValue) {
        // Closing out the active attendance record.
        checkoutCell.setValue(currentTime);
        return jsonResponse({
          status: "ok",
          action: "checkout",
          timestamp: isoTimestamp,
        });
      }

      // All previous records for today already closed → start a new entry.
      attendanceSheet.appendRow([
        uid,
        student.SUID,
        currentDate,
        currentTime,
        "",
        fullName,
      ]);
      return jsonResponse({
        status: "ok",
        action: "checkin",
        timestamp: isoTimestamp,
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
 * Handles GET requests to provide dashboard data.
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doGet() {
  const studentsSheet = getSheet(STUDENTS_SHEET_NAME);
  const attendanceSheet = getSheet(ATTENDANCE_SHEET_NAME);
  const unregisteredSheet = getSheet(UNREGISTERED_SHEET_NAME);

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
  const spreadsheet = SpreadsheetApp.openByName(CLOUDATTEND_DB);
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
  const department = (data.department || "").toString().trim();
  const year = (data.year || "").toString().trim();

  const missing = [];
  if (!cardUid) missing.push("cardUid");
  if (!suid) missing.push("suid");
  if (!firstName) missing.push("firstName");
  if (!lastName) missing.push("lastName");
  if (!department) missing.push("department");
  if (!year) missing.push("year");

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
    Department: department,
    Year: year,
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
  const expectedHeaders = [
    "CARD_UID",
    "SUID",
    "FirstName",
    "LastName",
    "Department",
    "Year",
  ];

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
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  output.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (statusCode) {
    output.setHeader("X-Status-Code", String(statusCode));
  }
  return output;
}
