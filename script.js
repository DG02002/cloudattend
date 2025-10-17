// CloudAttend Dashboard logic.
// Fetches data from the Apps Script endpoint and updates the UI with vanilla JS.

/**
 * @typedef {Record<string, any>} AnyRecord
 */

const APPS_SCRIPT_BASE_URL = "https://script.google.com/macros/s/";
const APPS_SCRIPT_DEPLOYMENT_ID =
  "AKfycbz0F26gZ5EX5VNtZYW2Tr_gyGVgcEMX0LkXSdf4Q64apiLkEBvbplifFICe1TgEHtTo";
const API_URL = `${APPS_SCRIPT_BASE_URL}${APPS_SCRIPT_DEPLOYMENT_ID}/exec`;
const LOG_PREFIX = "[CloudAttend]";

const ATTENDANCE_COLUMN_COUNT = 4;
const AUTO_REFRESH_INTERVAL_MS = 15000;
const AUTO_REFRESH_ERROR_WINDOW_MS = 30000;
const ADD_STUDENT_LABEL = "Add student";
const ADD_STUDENT_LOADING_LABEL = "Adding...";

/**
 * @type {{
 *   students: AnyRecord[];
 *   attendance: AnyRecord[];
 *   unregistered: AnyRecord[];
 *   todayKey: string;
 *   searchTerm: string;
 *   selectedCardUid: string;
 *   loading: boolean;
 * }}
 */
const state = {
  students: [],
  attendance: [],
  unregistered: [],
  todayKey: getTodayKey(),
  searchTerm: "",
  selectedCardUid: "",
  loading: false,
};

/** @type {number | undefined} */
let autoRefreshTimer;
/** @type {{ silent: boolean } | null} */
let pendingRefreshOptions = null;
let lastAutoRefreshErrorAt = 0;

const elements = {
  lastUpdated: /** @type {HTMLSpanElement | null} */ (
    document.getElementById("lastUpdated")
  ),
  searchInput: /** @type {HTMLInputElement} */ (
    document.getElementById("searchInput")
  ),
  attendanceBody: /** @type {HTMLTableSectionElement} */ (
    document.getElementById("attendanceBody")
  ),
  attendanceRegion: /** @type {HTMLElement | null} */ (
    document.querySelector('[data-busy-target="attendance"]')
  ),
  attendanceSpinner: /** @type {HTMLElement | null} */ (
    document.getElementById("attendanceSpinner")
  ),
  unregisteredList: /** @type {HTMLElement} */ (
    document.getElementById("unregisteredList")
  ),
  form: /** @type {HTMLFormElement} */ (
    document.getElementById("registrationForm")
  ),
  formSubmitButton: /** @type {HTMLButtonElement | null} */ (
    document.querySelector("#registrationForm button[type='submit']")
  ),
  formCardUid: /** @type {HTMLInputElement} */ (
    document.getElementById("formCardUid")
  ),
  formSuid: /** @type {HTMLInputElement} */ (
    document.getElementById("formSuid")
  ),
  formFirstName: /** @type {HTMLInputElement} */ (
    document.getElementById("formFirstName")
  ),
  formLastName: /** @type {HTMLInputElement} */ (
    document.getElementById("formLastName")
  ),
  clearSelectionButton: /** @type {HTMLButtonElement} */ (
    document.getElementById("clearSelectionButton")
  ),
  toast: /** @type {HTMLElement} */ (document.getElementById("toast")),
};

document.addEventListener("DOMContentLoaded", async () => {
  try {
    bindEventListeners();
    await loadDashboard();
  } catch (error) {
    console.error(LOG_PREFIX, "Initialization error:", error);
  } finally {
    startAutoRefresh();
  }
});

function bindEventListeners() {
  elements.searchInput.addEventListener("input", (event) => {
    const input = /** @type {HTMLInputElement} */ (event.currentTarget);
    state.searchTerm = (input.value || "").toLowerCase().trim();
    renderAttendance();
  });

  elements.form.addEventListener("submit", async (event) => {
    event.preventDefault();
    await submitRegistration();
  });

  elements.clearSelectionButton.addEventListener("click", () => {
    selectCard("");
  });
}

/**
 * @param {{ silent?: boolean }} [options]
 */
async function loadDashboard(options) {
  const config = {
    silent: false,
    ...(options || {}),
  };

  if (state.loading) {
    pendingRefreshOptions = config;
    return;
  }

  setLoading(true, { silent: config.silent });
  try {
    const response = await fetch(API_URL);

    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}`);
    }

    const payload = await response.json();
    if (!payload || payload.status !== "ok" || !payload.data) {
      throw new Error("Malformed response from API");
    }

    state.students = Array.isArray(payload.data.students)
      ? payload.data.students
      : [];
    state.attendance = Array.isArray(payload.data.attendance)
      ? payload.data.attendance
      : [];
    state.unregistered = Array.isArray(payload.data.unregisteredCards)
      ? payload.data.unregisteredCards
      : [];

    state.todayKey = getTodayKey();
    renderAttendance();
    renderUnregistered();
    updateTimestamp();
    if (!config.silent) {
      showToast("Data refreshed.", "success");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(LOG_PREFIX, "Dashboard load failed:", error);
    console.error(LOG_PREFIX, "Dashboard load error message:", message);
    console.error(
      LOG_PREFIX,
      "Dashboard load stack:",
      error instanceof Error ? error.stack : "N/A"
    );
    if (config.silent) {
      const now = Date.now();
      if (now - lastAutoRefreshErrorAt > AUTO_REFRESH_ERROR_WINDOW_MS) {
        showToast("Auto refresh failed. Check your connection.", "error");
        lastAutoRefreshErrorAt = now;
      }
    } else {
      renderErrorRows("We can't refresh data right now.");
      showToast("We can't refresh data right now.", "error");
    }
  } finally {
    setLoading(false, { silent: config.silent });
    if (pendingRefreshOptions) {
      const nextOptions = pendingRefreshOptions;
      pendingRefreshOptions = null;
      await loadDashboard(nextOptions);
    }
  }
}

/**
 * Renders all attendance records and metrics.
 */
function renderAttendance() {
  const attendanceBody = elements.attendanceBody;
  attendanceBody.innerHTML = "";

  /** @type {Map<string, AnyRecord>} */
  const studentLookup = new Map();
  state.students.forEach((student) => {
    const entry = /** @type {AnyRecord} */ (student);
    const key = (entry.CARD_UID || "").toString().trim().toUpperCase();
    if (key) {
      studentLookup.set(key, entry);
    }
  });
  const allRecords = state.attendance.filter((record) => !!record);

  const filtered = allRecords.filter((record) => {
    if (!state.searchTerm) {
      return true;
    }
    const name = getRecordName(record, studentLookup).toLowerCase();
    const suid = (record.SUID || "").toString().toLowerCase();
    return name.includes(state.searchTerm) || suid.includes(state.searchTerm);
  });

  // Sort by date (newest first), then by check-in time
  filtered.sort((a, b) => {
    const aDate = getRecordDate(a);
    const bDate = getRecordDate(b);
    if (aDate && bDate && aDate !== bDate) {
      return bDate.localeCompare(aDate);
    }
    if (!aDate && bDate) {
      return 1;
    }
    if (aDate && !bDate) {
      return -1;
    }
    const aKey = parseTimeToSortable(getCheckIn(a));
    const bKey = parseTimeToSortable(getCheckIn(b));
    return aKey - bKey;
  });

  if (filtered.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = ATTENDANCE_COLUMN_COUNT;
    cell.className = "empty";
    cell.textContent = allRecords.length
      ? "No matches for that search."
      : "No attendance records yet.";
    row.appendChild(cell);
    attendanceBody.appendChild(row);
    return;
  }

  const fragment = document.createDocumentFragment();
  filtered.forEach((record) => {
    const entry = /** @type {AnyRecord} */ (record);
    const row = document.createElement("tr");
    const name = getDisplayName(entry, studentLookup);
    const suid = entry.SUID || "";
    const checkInRaw = getCheckIn(entry);
    const checkOutRaw = getCheckOut(entry);
    const checkIn = formatTime12(checkInRaw) || "--";
    const checkOut = formatTime12(checkOutRaw) || "--";
    const isPresent = !checkOutRaw;

    if (isPresent) {
      row.classList.add("is-present");
    }

    row.appendChild(createCell(name));
    row.appendChild(createCell(suid));
    row.appendChild(createCell(checkIn));
    row.appendChild(createCell(checkOut));

    fragment.appendChild(row);
  });

  attendanceBody.appendChild(fragment);
}

function renderUnregistered() {
  const list = elements.unregisteredList;
  list.innerHTML = "";

  const pendingCards = state.unregistered.filter((rawEntry) => {
    const entry = /** @type {AnyRecord} */ (rawEntry);
    const status = (entry.Status || entry.status || "Pending")
      .toString()
      .toLowerCase();
    return status === "pending" || !entry.Status;
  });

  if (pendingCards.length === 0) {
    const message = document.createElement("p");
    message.className = "placeholder";
    message.textContent = "All cards are set up.";
    list.appendChild(message);
    return;
  }

  const fragment = document.createDocumentFragment();
  pendingCards.forEach((rawEntry) => {
    const entry = /** @type {AnyRecord} */ (rawEntry);
    const cardUid = (entry.CARD_UID || "").toString().trim().toUpperCase();
    const card = document.createElement("article");
    card.className = "pending-card";
    if (cardUid === state.selectedCardUid) {
      card.classList.add("is-selected");
    }

    const header = document.createElement("header");
    const title = document.createElement("strong");
    title.textContent = cardUid || "Unknown card";
    const status = document.createElement("span");
    status.textContent = "Pending";
    status.className = "badge";
    header.appendChild(title);
    header.appendChild(status);

    const meta = document.createElement("div");
    meta.className = "pending-meta";
    const dateValue = formatHumanDate(entry.Date);
    if (dateValue) {
      const date = document.createElement("span");
      date.textContent = dateValue;
      meta.appendChild(date);
    }
    const timeValue = formatTime12(entry.Time);
    if (timeValue) {
      const time = document.createElement("span");
      time.textContent = timeValue;
      meta.appendChild(time);
    }

    const linkBtn = document.createElement("button");
    linkBtn.className = "primary small";
    linkBtn.type = "button";
    linkBtn.textContent = "Link student";
    linkBtn.setAttribute(
      "aria-pressed",
      cardUid === state.selectedCardUid ? "true" : "false"
    );
    linkBtn.addEventListener("click", () => {
      selectCard(cardUid);
      elements.formSuid.focus();
    });

    card.appendChild(header);
    if (meta.childElementCount > 0) {
      card.appendChild(meta);
    }
    const actions = document.createElement("div");
    actions.className = "card-actions";
    actions.appendChild(linkBtn);
    card.appendChild(actions);
    fragment.appendChild(card);
  });

  list.appendChild(fragment);
}

/**
 * @param {string} cardUid
 */
function selectCard(cardUid) {
  state.selectedCardUid = cardUid;
  elements.formCardUid.value = cardUid || "";
  refreshPendingSelection();
}

function refreshPendingSelection() {
  const cards = elements.unregisteredList.querySelectorAll(".pending-card");
  cards.forEach((card) => {
    const title = card.querySelector("strong");
    const isSelected =
      title && title.textContent.trim().toUpperCase() === state.selectedCardUid;
    card.classList.toggle("is-selected", !!isSelected);
    const action = card.querySelector("button.primary");
    if (action) {
      action.setAttribute("aria-pressed", isSelected ? "true" : "false");
    }
  });
}

async function submitRegistration() {
  if (!elements.form.checkValidity()) {
    elements.form.reportValidity();
    return;
  }

  const payload = {
    action: "register",
    data: {
      cardUid: elements.formCardUid.value.trim(),
      suid: elements.formSuid.value.trim(),
      firstName: elements.formFirstName.value.trim(),
      lastName: elements.formLastName.value.trim(),
    },
  };

  try {
    setSubmitLoading(true);
    setFormDisabled(true);
    // Skip setting Content-Type so the browser treats this as a simple POST and avoids a failing CORS preflight.
    const response = await fetch(API_URL, {
      method: "POST",
      body: JSON.stringify(payload),
    });

    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result || result.status !== "ok") {
      throw new Error(result.message || "Registration failed");
    }

    showToast("Student linked.", "success");
    elements.form.reset();
    selectCard("");
    await loadDashboard();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(LOG_PREFIX, "Registration request failed:", message);
    showToast("We can't link that student right now.", "error");
  } finally {
    setFormDisabled(false);
    setSubmitLoading(false);
  }
}

/**
 * @param {boolean} isLoading
 * @param {{ silent?: boolean }} [options]
 */
function setLoading(isLoading, options) {
  const config = {
    silent: false,
    ...(options || {}),
  };

  state.loading = isLoading;

  if (elements.attendanceRegion) {
    elements.attendanceRegion.setAttribute(
      "aria-busy",
      isLoading ? "true" : "false"
    );
  }

  if (elements.attendanceSpinner) {
    elements.attendanceSpinner.classList.toggle("is-visible", isLoading);
  }
}

function startAutoRefresh() {
  stopAutoRefresh();
  autoRefreshTimer = window.setInterval(() => {
    void loadDashboard({ silent: true });
  }, AUTO_REFRESH_INTERVAL_MS);
}

function stopAutoRefresh() {
  if (autoRefreshTimer !== undefined) {
    window.clearInterval(autoRefreshTimer);
    autoRefreshTimer = undefined;
  }
}

/**
 * @param {boolean} disabled
 */
function setFormDisabled(disabled) {
  const controls = /** @type {Array<HTMLInputElement | HTMLButtonElement>} */ (
    Array.from(elements.form.querySelectorAll("input, button"))
  );
  controls.forEach((control) => {
    control.disabled = disabled;
  });
}

/**
 * Toggles the registration submit button's loading indicator.
 * @param {boolean} isLoading
 */
function setSubmitLoading(isLoading) {
  const button = elements.formSubmitButton;
  if (!button) {
    return;
  }
  button.classList.toggle("is-loading", isLoading);
  button.setAttribute("aria-busy", isLoading ? "true" : "false");
  button.textContent = isLoading
    ? ADD_STUDENT_LOADING_LABEL
    : ADD_STUDENT_LABEL;
}

function updateTimestamp() {
  const now = new Date();
  const formatter = new Intl.DateTimeFormat(undefined, {
    hour: "2-digit",
    minute: "2-digit",
  });
  if (elements.lastUpdated) {
    elements.lastUpdated.textContent = `Updated: ${formatter.format(now)}`;
  }
}

/**
 * @param {string} message
 */
function renderErrorRows(message) {
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = ATTENDANCE_COLUMN_COUNT;
  cell.className = "empty";
  cell.textContent = message;
  row.appendChild(cell);
  elements.attendanceBody.innerHTML = "";
  elements.attendanceBody.appendChild(row);
}

/**
 * @param {string} text
 * @param {string} [className]
 * @return {HTMLTableCellElement}
 */
function createCell(text, className) {
  const cell = document.createElement("td");
  cell.textContent = text || "--";
  if (className) {
    cell.className = className;
  }
  return cell;
}

/**
 * @param {AnyRecord} record
 * @param {Map<string, AnyRecord>} lookup
 * @return {string}
 */
function getRecordName(record, lookup) {
  const raw = (record.Name || record.name || "").trim();
  if (raw) {
    return raw;
  }
  const uid = (record.CARD_UID || "").toString().trim().toUpperCase();
  const student = lookup.get(uid);
  if (!student) {
    return "Unknown";
  }
  const first = student.FirstName || "";
  const last = student.LastName || "";
  return `${first} ${last}`.trim() || "Unknown";
}

// Attendance field fallbacks to handle header mismatches
/**
 * @param {AnyRecord} entry
 */
function getCheckIn(entry) {
  const value =
    entry.CheckInTime ||
    entry["Check In Time"] ||
    entry["Check-in"] ||
    entry.checkIn ||
    entry.checkin ||
    "";

  // If it's a Date object or timestamp string, format it
  if (value instanceof Date) {
    return formatTime12(value);
  }

  const str = (value || "").toString();
  // Check if it's an ISO timestamp (contains 'T' or 'Z')
  if (str.includes("T") || str.includes("Z")) {
    const date = new Date(str);
    if (!isNaN(date.getTime())) {
      return formatTime12(date);
    }
  }

  return str;
}

/**
 * @param {AnyRecord} entry
 */
function getCheckOut(entry) {
  const value =
    entry.CheckOutTime ||
    entry["Check Out Time"] ||
    entry["Check-out"] ||
    entry.checkout ||
    entry.checkOut ||
    "";

  // If it's a Date object or timestamp string, format it
  if (value instanceof Date) {
    return formatTime12(value);
  }

  const str = (value || "").toString();
  // Check if it's an ISO timestamp (contains 'T' or 'Z')
  if (str.includes("T") || str.includes("Z")) {
    const date = new Date(str);
    if (!isNaN(date.getTime())) {
      return formatTime12(date);
    }
  }

  return str;
}

/**
 * @param {AnyRecord} entry
 * @param {Map<string, AnyRecord>} lookup
 */
function getDisplayName(entry, lookup) {
  return (
    (entry.Name || entry.name || "").trim() || getRecordName(entry, lookup)
  );
}

/**
 * Extracts an ISO-like date (YYYY-MM-DD) from an attendance record when possible.
 * @param {AnyRecord} entry
 * @return {string}
 */
function getRecordDate(entry) {
  const fromKey = normalizeDateString(entry.DateKey || entry.dateKey);
  if (fromKey) {
    return fromKey;
  }

  const rawSource =
    entry.Date ||
    entry.date ||
    entry.CheckInDate ||
    entry["Check In Date"] ||
    entry.Timestamp ||
    entry.timestamp ||
    "";

  if (rawSource instanceof Date) {
    return formatDateYmd(rawSource);
  }

  const asString = (rawSource || "").toString();
  const normalized = normalizeDateString(asString);
  if (normalized) {
    return normalized;
  }

  const parsed = new Date(asString);
  if (!isNaN(parsed.getTime())) {
    return formatDateYmd(parsed);
  }

  return "";
}

/**
 * Normalizes the date portion of a string to YYYY-MM-DD when possible.
 * @param {unknown} value
 * @return {string}
 */
function normalizeDateString(value) {
  const raw = (value || "").toString().trim();
  if (!raw) {
    return "";
  }
  const isoCandidate = raw.slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoCandidate)) {
    return isoCandidate;
  }
  return "";
}

/**
 * Formats a Date into YYYY-MM-DD.
 * @param {Date} date
 * @return {string}
 */
function formatDateYmd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function getTodayKey() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

/** @type {number | undefined} */
let toastTimeout;

/**
 * @param {string} message
 * @param {"info" | "success" | "error"} [type]
 */
function showToast(message, type = "info") {
  if (!elements.toast) {
    return;
  }
  elements.toast.textContent = message;
  elements.toast.classList.remove("is-visible", "is-error", "is-success");
  if (type === "error") {
    elements.toast.classList.add("is-error");
  } else if (type === "success") {
    elements.toast.classList.add("is-success");
  }
  // Trigger reflow to restart animation when showing consecutively.
  void elements.toast.offsetWidth;
  elements.toast.classList.add("is-visible");

  if (toastTimeout !== undefined) {
    clearTimeout(toastTimeout);
  }
  toastTimeout = setTimeout(() => {
    elements.toast.classList.remove("is-visible");
  }, 3200);
}

/**
 * Converts a date string to human-readable form if possible.
 * Accepts already human-readable strings and returns them unchanged.
 * @param {string} value
 */
function formatHumanDate(value) {
  const raw = (value || "").toString().trim();
  if (!raw) {
    return "--";
  }

  const isoMatch = raw.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) {
    const dt = new Date(`${isoMatch[1]}T00:00:00Z`);
    if (!isNaN(dt.getTime())) {
      return dt.toLocaleDateString(undefined, {
        day: "numeric",
        month: "long",
        year: "numeric",
      });
    }
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return parsed.toLocaleDateString(undefined, {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  }

  return raw;
}

/**
 * Formats time into 12-hour with AM/PM when possible.
 * If value is already human-readable (has AM/PM), returns as-is.
 * @param {string | Date} value
 */
function formatTime12(value) {
  if (value instanceof Date) {
    return value.toLocaleTimeString(undefined, {
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    });
  }
  const raw = (value || "").toString().trim();
  if (!raw) return "";
  if (/am|pm/i.test(raw)) return raw.toUpperCase();
  if (raw.includes("T") || raw.includes("Z")) {
    const parsed = new Date(raw);
    if (!isNaN(parsed.getTime())) {
      return parsed.toLocaleTimeString(undefined, {
        hour: "numeric",
        minute: "2-digit",
        hour12: true,
      });
    }
  }
  // Expect HH:mm or HH:mm:ss
  const match = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return raw;
  let hours = parseInt(match[1], 10);
  const minutes = match[2];
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return `${hours}:${minutes} ${ampm}`;
}

/**
 * Parses a time string (either HH:mm[:ss] or h:mm AM/PM) into minutes since midnight.
 * Unparseable values are sent to end of list.
 * @param {string | Date} value
 * @return {number}
 */
function parseTimeToSortable(value) {
  if (value instanceof Date) {
    return value.getHours() * 60 + value.getMinutes();
  }
  const raw = (value || "").toString().trim();
  if (!raw) return Number.POSITIVE_INFINITY;
  // 12-hour format with AM/PM
  const m12 = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)$/i);
  if (m12) {
    let h = parseInt(m12[1], 10);
    const m = parseInt(m12[2], 10);
    const ampm = m12[4].toUpperCase();
    if (h === 12) h = 0;
    if (ampm === "PM") h += 12;
    return h * 60 + m;
  }
  // 24-hour format
  const m24 = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m24) {
    const h = parseInt(m24[1], 10);
    const m = parseInt(m24[2], 10);
    return h * 60 + m;
  }
  return Number.POSITIVE_INFINITY;
}

// Helpers to compute a dateKey when only a Date string is present
/**
 * @param {string} dateValue
 */
function getTodayKeyFallback(dateValue) {
  const raw = (dateValue || "").toString();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw.slice(0, 10))) {
    return raw.slice(0, 10);
  }
  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, "0");
    const d = String(parsed.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return state.todayKey;
}
