// CloudAttend Dashboard logic.
// Fetches data from the Apps Script endpoint and updates the UI with vanilla JS.

/**
 * @typedef {Record<string, any>} AnyRecord
 */

const API_URL =
  "https://script.google.com/macros/s/AKfycbyeu3vvlPiiu1a8tqThb8WSXtKytRGrGr6cfSVrLht8Iq3vNgTTxItfXaEPOuXxu8NG/exec";

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

const elements = {
  refreshButton: /** @type {HTMLButtonElement} */ (
    document.getElementById("refreshButton")
  ),
  lastUpdated: /** @type {HTMLSpanElement} */ (
    document.getElementById("lastUpdated")
  ),
  searchInput: /** @type {HTMLInputElement} */ (
    document.getElementById("searchInput")
  ),
  attendanceBody: /** @type {HTMLTableSectionElement} */ (
    document.getElementById("attendanceBody")
  ),
  presentCount: /** @type {HTMLSpanElement} */ (
    document.getElementById("presentCount")
  ),
  totalCount: /** @type {HTMLSpanElement} */ (
    document.getElementById("totalCount")
  ),
  unregisteredList: /** @type {HTMLElement} */ (
    document.getElementById("unregisteredList")
  ),
  form: /** @type {HTMLFormElement} */ (
    document.getElementById("registrationForm")
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
  formDepartment: /** @type {HTMLInputElement} */ (
    document.getElementById("formDepartment")
  ),
  formYear: /** @type {HTMLInputElement} */ (
    document.getElementById("formYear")
  ),
  clearSelectionButton: /** @type {HTMLButtonElement} */ (
    document.getElementById("clearSelectionButton")
  ),
  toast: /** @type {HTMLElement} */ (document.getElementById("toast")),
};

document.addEventListener("DOMContentLoaded", () => {
  bindEventListeners();
  loadDashboard();
});

function bindEventListeners() {
  elements.refreshButton.addEventListener("click", loadDashboard);

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

async function loadDashboard() {
  if (state.loading) {
    return;
  }

  setLoading(true);
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
    showToast("Dashboard updated", "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    renderErrorRows(message || "Unable to load data");
    showToast("Failed to fetch data", "error");
  } finally {
    setLoading(false);
  }
}

/**
 * Renders todays attendance list and metrics.
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

  const todayRecords = state.attendance.filter((record) => {
    const entry = /** @type {AnyRecord} */ (record);
    const recordDate = (entry.Date || entry.date || "").toString().slice(0, 10);
    return recordDate === state.todayKey;
  });

  const filtered = todayRecords.filter((record) => {
    if (!state.searchTerm) {
      return true;
    }
    const name = getRecordName(record, studentLookup).toLowerCase();
    const suid = (record.SUID || "").toString().toLowerCase();
    return name.includes(state.searchTerm) || suid.includes(state.searchTerm);
  });

  filtered.sort((a, b) => {
    const aTime = `${a.Date || ""} ${a.CheckInTime || ""}`;
    const bTime = `${b.Date || ""} ${b.CheckInTime || ""}`;
    return aTime.localeCompare(bTime);
  });

  const presentCount = filtered.reduce((count, record) => {
    return !record.CheckOutTime ? count + 1 : count;
  }, 0);

  elements.presentCount.textContent = presentCount.toString();
  elements.totalCount.textContent = filtered.length.toString();

  if (filtered.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 5;
    cell.className = "empty";
    cell.textContent = todayRecords.length
      ? "No results match your search"
      : "No attendance recorded today";
    row.appendChild(cell);
    attendanceBody.appendChild(row);
    return;
  }

  const fragment = document.createDocumentFragment();
  filtered.forEach((record) => {
    const entry = /** @type {AnyRecord} */ (record);
    const row = document.createElement("tr");
    const cardUid = (entry.CARD_UID || "").toString().trim().toUpperCase();
    const name = getRecordName(entry, studentLookup);
    const suid = entry.SUID || "";
    const checkIn = entry.CheckInTime || "--";
    const checkOut = entry.CheckOutTime || "--";
    const isPresent = !entry.CheckOutTime;

    if (isPresent) {
      row.classList.add("is-present");
    }

    row.appendChild(createCell(name));
    row.appendChild(createCell(suid, "align-right"));
    row.appendChild(createCell(checkIn));
    row.appendChild(createCell(checkOut));
    row.appendChild(createCell(isPresent ? "In" : "Complete"));

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
    message.textContent = "No unregistered cards";
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
    title.textContent = cardUid || "Unknown UID";
    const status = document.createElement("span");
    status.textContent = "Pending";
    status.className = "badge";
    header.appendChild(title);
    header.appendChild(status);

    const details = document.createElement("dl");
    details.appendChild(createDefinition("Date", entry.Date || "--"));
    details.appendChild(createDefinition("Time", entry.Time || "--"));

    const button = document.createElement("button");
    button.className = "primary";
    button.textContent = "Register";
    button.addEventListener("click", () => {
      selectCard(cardUid);
      elements.formSuid.focus();
    });

    card.appendChild(header);
    card.appendChild(details);
    card.appendChild(button);
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
      department: elements.formDepartment.value.trim(),
      year: elements.formYear.value.trim(),
    },
  };

  try {
    setFormDisabled(true);
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result || result.status !== "ok") {
      throw new Error(result.message || "Registration failed");
    }

    showToast("Student registered", "success");
    elements.form.reset();
    selectCard("");
    await loadDashboard();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    showToast(message || "Unable to register student", "error");
  } finally {
    setFormDisabled(false);
  }
}

/**
 * @param {boolean} isLoading
 */
function setLoading(isLoading) {
  state.loading = isLoading;
  elements.refreshButton.disabled = isLoading;
  if (isLoading) {
    elements.refreshButton.textContent = "Refreshing";
  } else {
    elements.refreshButton.textContent = "Refresh";
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

function updateTimestamp() {
  const now = new Date();
  const formatter = new Intl.DateTimeFormat(undefined, {
    hour: "2-digit",
    minute: "2-digit",
  });
  elements.lastUpdated.textContent = `Updated: ${formatter.format(now)}`;
}

/**
 * @param {string} message
 */
function renderErrorRows(message) {
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = 5;
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
 * @param {string} label
 * @param {string} value
 * @return {DocumentFragment}
 */
function createDefinition(label, value) {
  const term = document.createElement("dt");
  term.textContent = label;
  const description = document.createElement("dd");
  description.textContent = value || "--";
  const fragment = document.createDocumentFragment();
  fragment.appendChild(term);
  fragment.appendChild(description);
  return fragment;
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
