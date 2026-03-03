const STORAGE_KEY = "pixelbug-calendar-ui-v1";
const COLOR_OPTIONS = [
  "#ee8f8f",
  "#f2b779",
  "#e6cf72",
  "#9fd67b",
  "#73cfc0",
  "#79afea",
  "#9b92ef",
  "#c78eeb",
  "#ee97bf",
  "#8fc3a0",
  "#d9b27c",
  "#7fd6f5",
  "#f59fc1"
];

const weekdayLabels = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const SHOW_BLOCK_MINUTES = 120;
const DAY_START_HOUR = 6;
const DAY_END_HOUR = 24;

function uid(prefix) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function formatCurrency(amount) {
  if (amount === "" || amount === null || amount === undefined) return "-";
  const value = Number(amount);
  if (Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0
  }).format(value);
}

function formatDate(dateStr) {
  if (!dateStr) return "-";
  const date = new Date(`${dateStr}T00:00:00`);
  return date.toLocaleDateString("en-IN", {
    day: "numeric",
    month: "short",
    year: "numeric"
  });
}

function formatDateRange(fromDate, toDate) {
  if (!fromDate && !toDate) return "-";
  if (!toDate || fromDate === toDate) return formatDate(fromDate);
  return `${formatDate(fromDate)} - ${formatDate(toDate)}`;
}

function monthLabel(year, month) {
  return new Date(year, month, 1).toLocaleDateString("en-IN", {
    month: "long",
    year: "numeric"
  });
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseDateKey(value) {
  if (!value) return new Date();
  return new Date(`${value}T00:00:00`);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function startOfWeek(date) {
  return addDays(date, -date.getDay());
}

function formatShortDate(date) {
  return date.toLocaleDateString("en-IN", {
    day: "numeric",
    month: "short"
  });
}

function formatWeekdayDate(date) {
  return date.toLocaleDateString("en-IN", {
    weekday: "short",
    day: "numeric",
    month: "short"
  });
}

function timeLabel(hour) {
  const date = new Date();
  date.setHours(hour, 0, 0, 0);
  return date.toLocaleTimeString("en-IN", {
    hour: "numeric",
    hour12: true
  });
}

function parseTimeToMinutes(time) {
  if (!time) return 12 * 60;
  const [hours, minutes] = time.split(":").map(Number);
  if (Number.isNaN(hours) || Number.isNaN(minutes)) return 12 * 60;
  return (hours * 60) + minutes;
}

function getFocusDate() {
  return parseDateKey(state.view.focusDate);
}

function syncViewDateParts(date) {
  state.view.focusDate = dateKey(date);
  state.view.year = date.getFullYear();
  state.view.month = date.getMonth();
}

function ensureViewState() {
  const today = new Date();
  if (!state.view) {
    state.view = {
      year: today.getFullYear(),
      month: today.getMonth(),
      mode: "month",
      focusDate: dateKey(today)
    };
  }

  if (!state.view.mode) {
    state.view.mode = "month";
  }

  if (!state.view.focusDate) {
    syncViewDateParts(new Date(state.view.year ?? today.getFullYear(), state.view.month ?? today.getMonth(), today.getDate()));
  }
}

function getVisibleRangeLabel() {
  const focusDate = getFocusDate();
  if (state.view.mode === "day") {
    return formatWeekdayDate(focusDate);
  }

  if (state.view.mode === "week") {
    const start = startOfWeek(focusDate);
    const end = addDays(start, 6);
    return `${formatShortDate(start)} - ${formatShortDate(end)}`;
  }

  return monthLabel(state.view.year, state.view.month);
}

function getShowsForDate(shows, value) {
  return shows.filter((show) => isDateWithinCalendarBlock(show, value));
}

function getShowStartDate(show) {
  return show.showDateFrom || show.showDate || "";
}

function getShowEndDate(show) {
  return show.showDateTo || show.showDateFrom || show.showDate || "";
}

function getCalendarBlockStartDate(show) {
  const onwardDates = (show.assignments || [])
    .map((assignment) => assignment.onwardTravelDate)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  return onwardDates[0] || getShowStartDate(show);
}

function getCalendarBlockEndDate(show) {
  return getShowEndDate(show);
}

function isDateWithinCalendarBlock(show, value) {
  const start = getCalendarBlockStartDate(show);
  const end = getCalendarBlockEndDate(show);
  if (!start) return false;
  return value >= start && value <= end;
}

function isDateWithinShowRange(show, value) {
  const start = getShowStartDate(show);
  const end = getShowEndDate(show);
  if (!start) return false;
  return value >= start && value <= end;
}

function monthGroupLabel(dateStr) {
  const date = parseDateKey(dateStr);
  return date.toLocaleDateString("en-IN", {
    month: "long",
    year: "numeric"
  });
}

function getPrimaryCrewName(show) {
  const crewNames = show.assignments
    .map((assignment) => getUserById(assignment.crewId)?.name)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  return crewNames[0] || "";
}

function sortShows(shows, mode) {
  const items = [...shows];

  if (mode === "crew") {
    return items.sort((a, b) => getPrimaryCrewName(a).localeCompare(getPrimaryCrewName(b)) || getShowStartDate(a).localeCompare(getShowStartDate(b)));
  }

  if (mode === "status") {
    const rank = (show) => (show.showStatus === "tentative" ? 1 : 0);
    return items.sort((a, b) => rank(a) - rank(b) || getShowStartDate(a).localeCompare(getShowStartDate(b)));
  }

  if (mode === "city") {
    return items.sort((a, b) => (a.location || "").localeCompare(b.location || "") || getShowStartDate(a).localeCompare(getShowStartDate(b)));
  }

  if (mode === "client") {
    return items.sort((a, b) => (a.client || "").localeCompare(b.client || "") || getShowStartDate(a).localeCompare(getShowStartDate(b)));
  }

  return items.sort((a, b) => getShowStartDate(a).localeCompare(getShowStartDate(b)) || (a.showTime || "").localeCompare(b.showTime || ""));
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function excelColumnName(index) {
  let column = "";
  let value = index;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    value = Math.floor((value - 1) / 26);
  }
  return column;
}

function makeWorksheetCell(cellRef, value, type = "inlineStr", styleId = null) {
  if (type === "n") {
    return `<c r="${cellRef}"${styleId ? ` s="${styleId}"` : ""}><v>${Number(value || 0)}</v></c>`;
  }

  return `<c r="${cellRef}" t="inlineStr"${styleId ? ` s="${styleId}"` : ""}><is><t>${escapeXml(value)}</t></is></c>`;
}

function buildShowsSheetXml(shows) {
  const headers = [
    "Show Date",
    "No. of days",
    "Show Name",
    "Client",
    "Location",
    "Amount of the Show",
    "Assigned Crew",
    "Operator Amount"
  ];

  const rows = [];
  const merges = [];

  rows.push(`<row r="1">${headers.map((header, index) => makeWorksheetCell(`${excelColumnName(index + 1)}1`, header, "inlineStr", 1)).join("")}</row>`);

  let rowNumber = 2;
  shows.forEach((show) => {
    const assignments = show.assignments.length ? show.assignments : [{ crewId: "", operatorAmount: "" }];
    const startRow = rowNumber;

    assignments.forEach((assignment, assignmentIndex) => {
      const cells = [];
      const crewUser = assignment.crewId ? getUserById(assignment.crewId) : null;
      const startDate = getShowStartDate(show);
      const endDate = getShowEndDate(show);
      const numberOfDays = startDate && endDate
        ? Math.max(1, Math.round((parseDateKey(endDate) - parseDateKey(startDate)) / (1000 * 60 * 60 * 24)) + 1)
        : 1;

      if (assignmentIndex === 0) {
        cells.push(makeWorksheetCell(`A${rowNumber}`, endDate || startDate || ""));
        cells.push(makeWorksheetCell(`B${rowNumber}`, numberOfDays, "n"));
        cells.push(makeWorksheetCell(`C${rowNumber}`, show.showName));
        cells.push(makeWorksheetCell(`D${rowNumber}`, show.client));
        cells.push(makeWorksheetCell(`E${rowNumber}`, show.location));
        cells.push(makeWorksheetCell(`F${rowNumber}`, show.amountShow, "n"));
      }

      cells.push(makeWorksheetCell(`G${rowNumber}`, crewUser?.name || ""));
      cells.push(makeWorksheetCell(`H${rowNumber}`, assignment.operatorAmount, "n"));
      rows.push(`<row r="${rowNumber}">${cells.join("")}</row>`);
      rowNumber += 1;
    });

    const endRow = rowNumber - 1;
    if (endRow > startRow) {
      ["A", "B", "C", "D", "E", "F"].forEach((column) => merges.push(`${column}${startRow}:${column}${endRow}`));
    }
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="14" customWidth="1"/>
    <col min="2" max="2" width="12" customWidth="1"/>
    <col min="3" max="3" width="24" customWidth="1"/>
    <col min="4" max="4" width="24" customWidth="1"/>
    <col min="5" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="18" customWidth="1"/>
    <col min="7" max="7" width="22" customWidth="1"/>
    <col min="8" max="8" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    ${rows.join("")}
  </sheetData>
  ${merges.length ? `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>` : ""}
</worksheet>`;
}

function crc32(bytes) {
  let crc = -1;
  for (let i = 0; i < bytes.length; i += 1) {
    crc ^= bytes[i];
    for (let j = 0; j < 8; j += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ -1) >>> 0;
}

function createZip(files) {
  const encoder = new TextEncoder();
  const fileEntries = files.map((file) => ({
    nameBytes: encoder.encode(file.name),
    dataBytes: encoder.encode(file.data)
  }));

  let localOffset = 0;
  const localParts = [];
  const centralParts = [];

  fileEntries.forEach((file) => {
    const checksum = crc32(file.dataBytes);
    const localHeader = new Uint8Array(30 + file.nameBytes.length + file.dataBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, 0, true);
    localView.setUint16(12, 0, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, file.dataBytes.length, true);
    localView.setUint32(22, file.dataBytes.length, true);
    localView.setUint16(26, file.nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(file.nameBytes, 30);
    localHeader.set(file.dataBytes, 30 + file.nameBytes.length);
    localParts.push(localHeader);

    const centralHeader = new Uint8Array(46 + file.nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 0, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, file.dataBytes.length, true);
    centralView.setUint32(24, file.dataBytes.length, true);
    centralView.setUint16(28, file.nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, localOffset, true);
    centralHeader.set(file.nameBytes, 46);
    centralParts.push(centralHeader);

    localOffset += localHeader.length;
  });

  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, fileEntries.length, true);
  endView.setUint16(10, fileEntries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, localOffset, true);
  endView.setUint16(20, 0, true);

  return new Blob([...localParts, ...centralParts, endRecord], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function exportShowsMonthExcel(monthKey, shows) {
  if (!monthKey || !shows.length) return;

  const files = [
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
    },
    {
      name: "_rels/.rels",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    },
    {
      name: "xl/workbook.xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${escapeXml(monthKey)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    },
    {
      name: "xl/styles.xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="10"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><name val="Arial"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE9EEF5"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`
    },
    {
      name: "xl/worksheets/sheet1.xml",
      data: buildShowsSheetXml(shows)
    }
  ];

  const blob = createZip(files);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `pixelbug-${monthKey}-shows.xlsx`;
  link.click();
  URL.revokeObjectURL(link.href);
}

function seedState() {
  return {
    users: [],
    shows: [],
    currentUserId: null,
    view: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(),
      mode: "month",
      focusDate: dateKey(new Date())
    },
    ui: {
      editingShowId: null,
      activeSidebarTab: "calendarPanel",
      activeShowMonth: null,
      selectedShowYear: "all",
      showSortMode: "date",
      selectedCrewFilter: "all",
      expandedCalendarWeeks: {},
      selectedCalendarShowId: null,
      calendarReturnMode: "month",
      authPanelMode: "profile",
      authPanelOpen: false
    }
  };
}

function loadLocalUiState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return seedState();
  }

  try {
    const parsed = JSON.parse(raw);
    return {
      ...seedState(),
      view: parsed.view || seedState().view,
      ui: parsed.ui || seedState().ui
    };
  } catch (error) {
    return seedState();
  }
}

function normalizeState() {
  ensureUiState();
  ensureViewState();

  state.shows = state.shows.map((show) => {
    const legacyTravelDate = show.travelDate || "";
    const legacyTravelSector = show.travelSector || "";
    const legacyTravelNotes = show.travelNotes || "";

    const assignments = (show.assignments || []).map((assignment) => ({
      ...assignment,
      onwardTravelDate: assignment.onwardTravelDate ?? assignment.travelDate ?? legacyTravelDate,
      returnTravelDate: assignment.returnTravelDate ?? "",
      onwardTravelSector: assignment.onwardTravelSector ?? assignment.travelSector ?? legacyTravelSector,
      returnTravelSector: assignment.returnTravelSector ?? "",
      notes: assignment.notes ?? assignment.travelNotes ?? legacyTravelNotes,
      travelDate: undefined,
      travelSector: undefined,
      travelNotes: undefined
    }));

    return {
      ...show,
      showDateFrom: show.showDateFrom ?? show.showDate ?? "",
      showDateTo: show.showDateTo ?? show.showDateFrom ?? show.showDate ?? "",
      showDate: show.showDateFrom ?? show.showDate ?? "",
      showStatus: show.showStatus === "tentative" ? "tentative" : "confirmed",
      assignments,
      travelDate: undefined,
      travelSector: undefined,
      travelNotes: undefined
    };
  });
}

function saveState(nextState) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    view: nextState.view,
    ui: nextState.ui
  }));
}

function applyServerState(payload) {
  state.users = payload.users || [];
  state.shows = payload.shows || [];
  state.currentUserId = payload.currentUserId || null;
  normalizeState();
}

async function apiRequest(path, options = {}) {
  const response = await fetch(path, {
    method: options.method || "GET",
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {})
    },
    credentials: "same-origin",
    body: options.body
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || "Request failed.");
  }
  return payload;
}

async function refreshFromServer() {
  const payload = await apiRequest("/api/bootstrap");
  applyServerState(payload);
}

async function syncAdminState() {
  const payload = await apiRequest("/api/admin/state", {
    method: "POST",
    body: JSON.stringify({
      users: state.users,
      shows: state.shows
    })
  });
  applyServerState(payload);
}

let state = loadLocalUiState();

function getCurrentUser() {
  return state.users.find((user) => user.id === state.currentUserId) || null;
}

function isAdmin(user) {
  return user?.role === "admin";
}

function getCrewUsers() {
  return state.users.filter((user) => (user.role === "crew" || user.role === "admin") && user.approved);
}

function getPendingUsers() {
  return state.users.filter((user) => !user.approved);
}

function getApprovedCrewOnlyUsers() {
  return state.users.filter((user) => user.role === "crew" && user.approved);
}

function getApprovedAdminUsers() {
  return state.users.filter((user) => user.role === "admin" && user.approved);
}

function getUserById(id) {
  return state.users.find((user) => user.id === id);
}

function ensureUiState() {
  if (!state.ui) {
    state.ui = {
      editingShowId: null,
      activeSidebarTab: "calendarPanel",
      activeShowMonth: null,
      selectedShowYear: "all",
      showSortMode: "date",
      selectedCrewFilter: "all",
      expandedCalendarWeeks: {},
      selectedCalendarShowId: null,
      calendarReturnMode: "month",
      authPanelMode: "profile",
      authPanelOpen: false
    };
  }

  if (!state.ui.activeSidebarTab) {
    state.ui.activeSidebarTab = "calendarPanel";
  }

  if (!state.ui.showSortMode) {
    state.ui.showSortMode = "date";
  }

  if (!state.ui.selectedShowYear) {
    state.ui.selectedShowYear = "all";
  }

  if (!state.ui.selectedCrewFilter) {
    state.ui.selectedCrewFilter = "all";
  }

  if (!state.ui.expandedCalendarWeeks || typeof state.ui.expandedCalendarWeeks !== "object") {
    state.ui.expandedCalendarWeeks = {};
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "selectedCalendarShowId")) {
    state.ui.selectedCalendarShowId = null;
  }

  if (!state.ui.calendarReturnMode) {
    state.ui.calendarReturnMode = "month";
  }

  if (!state.ui.authPanelMode) {
    state.ui.authPanelMode = "profile";
  }

  if (typeof state.ui.authPanelOpen !== "boolean") {
    state.ui.authPanelOpen = false;
  }
}

function resetEditingState() {
  ensureUiState();
  state.ui.editingShowId = null;
}

async function copyText(text) {
  if (!text) return false;
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return true;
  }
  return false;
}

function getUrlToken(name) {
  return new URL(window.location.href).searchParams.get(name);
}

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.remove("hidden");
  toast.classList.add("visible");
  window.clearTimeout(showToast.timeoutId);
  showToast.timeoutId = window.setTimeout(() => {
    toast.classList.remove("visible");
    toast.classList.add("hidden");
  }, 2400);
}

function getSidebarTabs(user) {
  return user && isAdmin(user)
    ? [
        { id: "calendarPanel", label: "Calendar", meta: "Month, week, day" },
        { id: "showFormPanel", label: "Create Show", meta: "Add or edit entries" },
        { id: "showsPanel", label: "All Shows", meta: "All scheduled shows" },
        { id: "crewAdminPanel", label: "Crew Management", meta: "Add or remove crew" }
      ]
    : [
        { id: "calendarPanel", label: "Calendar", meta: "Shared schedule" },
        { id: "showsPanel", label: "Shows", meta: "Visible entries" },
        ...(user?.role === "viewer" ? [{ id: "legendPanel", label: "Crew", meta: "Colors and teams" }] : [])
      ];
}

function ensureActiveSidebarTab(user) {
  const tabs = getSidebarTabs(user);
  if (!tabs.some((tab) => tab.id === state.ui.activeSidebarTab)) {
    state.ui.activeSidebarTab = tabs[0].id;
  }
}

function getTakenColors(excludeUserId = null) {
  return state.users
    .filter((user) => (user.role === "crew" || user.role === "admin") && user.approved && user.id !== excludeUserId)
    .map((user) => user.color)
    .filter(Boolean);
}

function visibleShowsForUser(user) {
  if (!user) return [];
  if (user.role === "admin" || user.role === "viewer") {
    return [...state.shows].sort((a, b) => getShowStartDate(a).localeCompare(getShowStartDate(b)));
  }

  return state.shows
    .filter((show) => show.assignments.some((assignment) => assignment.crewId === user.id))
    .sort((a, b) => getShowStartDate(a).localeCompare(getShowStartDate(b)));
}

function canSeeOperatorAmount(user, assignment) {
  return isAdmin(user) || assignment.crewId === user?.id;
}

function filterShowsBySelectedCrew(shows) {
  if (!state.ui.selectedCrewFilter || state.ui.selectedCrewFilter === "all") {
    return shows;
  }

  if (state.ui.selectedCrewFilter === "unassigned") {
    return shows.filter((show) => !show.assignments.length);
  }

  return shows.filter((show) => show.assignments.some((assignment) => assignment.crewId === state.ui.selectedCrewFilter));
}

function renderCrewFilterControl(selectId, options = {}) {
  const label = options.label ?? "Crew";
  const wrapperClass = options.compact ? "sort-control compact-control" : "sort-control";
  const allLabel = options.allLabel ?? "All Crew";
  const includeUnassigned = options.includeUnassigned ?? false;
  const crewUsers = getCrewUsers();
  return `
    <label class="${wrapperClass}">
      ${label ? `<span>${label}</span>` : ""}
      <select id="${selectId}">
        <option value="all" ${state.ui.selectedCrewFilter === "all" ? "selected" : ""}>${allLabel}</option>
        ${includeUnassigned ? `<option value="unassigned" ${state.ui.selectedCrewFilter === "unassigned" ? "selected" : ""}>Unassigned</option>` : ""}
        ${crewUsers.map((crewUser) => `<option value="${crewUser.id}" ${state.ui.selectedCrewFilter === crewUser.id ? "selected" : ""}>${crewUser.name}</option>`).join("")}
      </select>
    </label>
  `;
}

function renderAuthPanel() {
  const authPanel = document.getElementById("authPanel");
  const user = getCurrentUser();
  authPanel.innerHTML = "";

  const hasAdmin = state.users.some((item) => item.role === "admin" && item.approved);

  if (!hasAdmin) {
    authPanel.append(document.getElementById("adminSetupTemplate").content.cloneNode(true));
    renderColorChoices(null, "adminColorChoices");
    wireAdminSetupForm();
    return;
  }

  if (!user) {
    authPanel.classList.remove("hidden");
    if (getUrlToken("reset")) {
      authPanel.append(document.getElementById("resetPasswordTemplate").content.cloneNode(true));
      wireResetPasswordForm();
      return;
    }
    authPanel.append(document.getElementById("loginTemplate").content.cloneNode(true));
    wireAuthForms();
    return;
  }

  authPanel.innerHTML = "";
  authPanel.classList.add("hidden");
}

function wireAdminSetupForm() {
  const form = document.getElementById("adminSetupForm");
  if (!form) return;

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const email = data.get("email").toString().trim().toLowerCase();
    const message = document.getElementById("adminSetupMessage");
    const selectedColorButton = document.querySelector("#adminColorChoices .color-option.selected");

    if (state.users.some((user) => user.email.toLowerCase() === email)) {
      message.textContent = "That email already exists.";
      return;
    }

    if (!selectedColorButton) {
      message.textContent = "Choose a crew color for the admin account.";
      return;
    }

    try {
      const payload = await apiRequest("/api/setup-admin", {
        method: "POST",
        body: JSON.stringify({
          name: data.get("name").toString().trim(),
          email,
          phone: data.get("phone").toString().trim(),
          password: data.get("password").toString(),
          color: selectedColorButton.dataset.color
        })
      });

      applyServerState(payload);
      saveState(state);
      render();
      showToast("Admin account created.");
    } catch (error) {
      message.textContent = error.message;
    }
  });
}

function renderSidebarTabs() {
  const user = getCurrentUser();
  const nodes = [
    document.getElementById("sidebarTabsDesktop"),
    document.getElementById("sidebarTabsMobile")
  ].filter(Boolean);
  if (!nodes.length) return;
  if (!user) {
    nodes.forEach((node) => {
      node.innerHTML = "";
    });
    return;
  }
  ensureActiveSidebarTab(user);
  const tabs = getSidebarTabs(user);
  const markup = tabs.map((tab) => `
    <button type="button" class="sidebar-tab ${state.ui.activeSidebarTab === tab.id ? "active" : ""}" data-target-panel="${tab.id}">
      <strong>${tab.label}</strong>
      <span>${tab.meta}</span>
    </button>
  `).join("");

  nodes.forEach((node) => {
    node.innerHTML = markup;
    node.querySelectorAll("[data-target-panel]").forEach((button) => {
      button.addEventListener("click", () => {
        state.ui.activeSidebarTab = button.dataset.targetPanel;
        saveState(state);
        renderSidebarTabs();
        renderDashboard();
      });
    });
  });
}

function wireAuthForms() {
  const showRegister = document.getElementById("showRegister");
  const showLogin = document.getElementById("showLogin");
  const showForgotPassword = document.getElementById("showForgotPassword");
  const loginForm = document.getElementById("loginForm");

  if (showRegister) {
    showRegister.addEventListener("click", () => {
      const authPanel = document.getElementById("authPanel");
      authPanel.innerHTML = "";
      authPanel.append(document.getElementById("registerTemplate").content.cloneNode(true));
      renderColorChoices();
      wireRegisterForm();
    });
  }

  if (showLogin && !loginForm) {
    showLogin.addEventListener("click", () => {
      renderAuthPanel();
    });
  }

  if (showForgotPassword) {
    showForgotPassword.addEventListener("click", () => {
      const authPanel = document.getElementById("authPanel");
      authPanel.innerHTML = "";
      authPanel.append(document.getElementById("forgotPasswordTemplate").content.cloneNode(true));
      wireForgotPasswordForm();
    });
  }

  if (loginForm) {
    loginForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(loginForm);
      const email = form.get("email").toString().trim().toLowerCase();
      const password = form.get("password").toString();
      const message = document.getElementById("authMessage");
      message.textContent = "";

      try {
        const payload = await apiRequest("/api/login", {
          method: "POST",
          body: JSON.stringify({ email, password })
        });
        applyServerState(payload);
        saveState(state);
        render();
      } catch (error) {
        message.textContent = error.message;
      }
    });
  }
}

function wireForgotPasswordForm() {
  const form = document.getElementById("forgotPasswordForm");
  const backButton = document.getElementById("backToLoginFromReset");
  if (backButton) {
    backButton.addEventListener("click", () => {
      renderAuthPanel();
    });
  }
  if (!form) return;

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const message = document.getElementById("forgotPasswordMessage");
    message.textContent = "";

    try {
      const payload = await apiRequest("/api/request-password-reset", {
        method: "POST",
        body: JSON.stringify({
          email: data.get("email").toString().trim().toLowerCase()
        })
      });
      if (payload.resetUrl) {
        message.innerHTML = `Reset link ready. <a href="${payload.resetUrl}">Open reset link</a>`;
      } else {
        message.textContent = "If the account exists, a reset link has been sent.";
      }
    } catch (error) {
      message.textContent = error.message;
    }
  });
}

function wireResetPasswordForm() {
  const form = document.getElementById("resetPasswordForm");
  const backButton = document.getElementById("backToLoginAfterPasswordReset");
  if (backButton) {
    backButton.addEventListener("click", () => {
      const url = new URL(window.location.href);
      url.searchParams.delete("reset");
      window.history.replaceState({}, document.title, url.pathname + (url.search ? url.search : ""));
      renderAuthPanel();
    });
  }
  if (!form) return;

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const message = document.getElementById("resetPasswordMessage");
    const newPassword = data.get("newPassword").toString();
    const confirmPassword = data.get("confirmPassword").toString();
    if (newPassword !== confirmPassword) {
      message.textContent = "New passwords do not match.";
      return;
    }

    try {
      await apiRequest("/api/reset-password", {
        method: "POST",
        body: JSON.stringify({
          token: getUrlToken("reset"),
          newPassword
        })
      });
      const url = new URL(window.location.href);
      url.searchParams.delete("reset");
      window.history.replaceState({}, document.title, url.pathname + (url.search ? url.search : ""));
      renderAuthPanel();
      document.getElementById("authMessage").textContent = "Password reset complete. Sign in with the new password.";
    } catch (error) {
      message.textContent = error.message;
    }
  });
}

function wireRegisterForm() {
  const roleSelect = document.getElementById("roleSelect");
  const colorField = document.getElementById("colorField");
  const registerForm = document.getElementById("registerForm");
  const showLogin = document.getElementById("showLogin");

  const syncColorVisibility = () => {
    colorField.classList.toggle("hidden", !["crew", "admin"].includes(roleSelect.value));
  };

  roleSelect.addEventListener("change", syncColorVisibility);
  syncColorVisibility();

  showLogin.addEventListener("click", () => {
    renderAuthPanel();
  });

  registerForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = new FormData(registerForm);
    const role = form.get("role").toString();
    const email = form.get("email").toString().trim().toLowerCase();
    const message = document.getElementById("registerMessage");
    const selectedColorButton = document.querySelector(".color-option.selected");

    if (state.users.some((user) => user.email.toLowerCase() === email)) {
      message.textContent = "That email already exists.";
      return;
    }

    if (["crew", "admin"].includes(role) && !selectedColorButton) {
      message.textContent = "Choose an available crew color.";
      return;
    }

    try {
      const payload = await apiRequest("/api/register", {
        method: "POST",
        body: JSON.stringify({
          name: form.get("name").toString().trim(),
          email,
          phone: form.get("phone").toString().trim(),
          password: form.get("password").toString(),
          role,
          color: ["crew", "admin"].includes(role) ? selectedColorButton.dataset.color : null
        })
      });
      await refreshFromServer();
      saveState(state);
      renderAuthPanel();
      const authMessage = document.getElementById("authMessage");
      authMessage.textContent = "Account request submitted. Wait for admin approval, then sign in with your email and password.";
    } catch (error) {
      message.textContent = error.message;
    }
  });
}

function renderColorChoices(selected = null, containerId = "colorChoices", excludeUserId = null) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const taken = getTakenColors(excludeUserId);
  container.innerHTML = "";

  COLOR_OPTIONS.forEach((color) => {
    if (taken.includes(color) && selected !== color) {
      return;
    }
    const button = document.createElement("button");
    button.type = "button";
    button.className = `color-option ${selected === color ? "selected" : ""}`;
    button.style.background = color;
    button.dataset.color = color;
    button.title = color;
    button.addEventListener("click", () => {
      container.querySelectorAll(".color-option").forEach((node) => node.classList.remove("selected"));
      button.classList.add("selected");
    });
    container.append(button);
  });
}

function renderSessionActions(user) {
  const node = document.getElementById("sessionActions");
  node.innerHTML = "";
  if (!user) return;

  node.innerHTML = `
    <div class="profile-menu-wrap">
      <button type="button" class="ghost small ${state.ui.authPanelOpen ? "is-active" : ""}" id="profileMenuButton">Hey, ${user.name}</button>
      ${state.ui.authPanelOpen ? `
        <div class="profile-menu-panel">
          <div class="profile-menu-tabs">
            <button type="button" class="ghost small ${state.ui.authPanelMode === "profile" ? "is-active" : ""}" id="profilePanelButton">Info</button>
            <button type="button" class="ghost small ${state.ui.authPanelMode === "password" ? "is-active" : ""}" id="passwordPanelButton">Password</button>
          </div>
          ${state.ui.authPanelMode === "password" ? `
            <form id="changePasswordForm" class="stack tight profile-menu-form">
              <label>
                <span>Current Password</span>
                <input type="password" name="currentPassword" required>
              </label>
              <label>
                <span>New Password</span>
                <input type="password" name="newPassword" minlength="8" required>
              </label>
              <label>
                <span>Confirm New Password</span>
                <input type="password" name="confirmPassword" minlength="8" required>
              </label>
              <p class="muted-note">Use 8+ characters with uppercase, lowercase, and a number.</p>
              <button type="submit" class="secondary">Update Password</button>
              <div id="changePasswordMessage" class="message"></div>
            </form>
          ` : `
            <div class="stack tight profile-menu-info">
              <div><strong>${user.name}</strong></div>
              <div class="meta-line">${user.email}</div>
              <div class="meta-line">Role: ${user.role === "viewer" ? "View Only" : user.role === "admin" ? "Admin" : "Crew"}</div>
              <div class="meta-line">Phone: ${user.phone}</div>
            </div>
          `}
          <button type="button" class="secondary small" id="logoutButton">Logout</button>
        </div>
      ` : ""}
    </div>
  `;

  document.getElementById("profileMenuButton")?.addEventListener("click", () => {
    state.ui.authPanelOpen = !state.ui.authPanelOpen;
    saveState(state);
    renderSessionActions(user);
  });

  document.getElementById("profilePanelButton")?.addEventListener("click", () => {
    state.ui.authPanelMode = "profile";
    state.ui.authPanelOpen = true;
    saveState(state);
    renderSessionActions(user);
  });

  document.getElementById("passwordPanelButton")?.addEventListener("click", () => {
    state.ui.authPanelMode = "password";
    state.ui.authPanelOpen = true;
    saveState(state);
    renderSessionActions(user);
  });

  const changePasswordForm = document.getElementById("changePasswordForm");
  if (changePasswordForm) {
    changePasswordForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(changePasswordForm);
      const message = document.getElementById("changePasswordMessage");
      const currentPassword = form.get("currentPassword").toString();
      const newPassword = form.get("newPassword").toString();
      const confirmPassword = form.get("confirmPassword").toString();

      if (newPassword !== confirmPassword) {
        message.textContent = "New passwords do not match.";
        return;
      }

      try {
        await apiRequest("/api/change-password", {
          method: "POST",
          body: JSON.stringify({ currentPassword, newPassword })
        });
        message.textContent = "Password updated.";
        changePasswordForm.reset();
      } catch (error) {
        message.textContent = error.message;
      }
    });
  }

  document.getElementById("logoutButton")?.addEventListener("click", () => {
    apiRequest("/api/logout", { method: "POST" })
      .then(() => {
        state.currentUserId = null;
        state.ui.authPanelOpen = false;
        saveState(state);
        render();
      })
      .catch((error) => showToast(error.message));
  });
}

function renderDashboard() {
  const dashboard = document.getElementById("dashboard");
  const user = getCurrentUser();
  const title = document.getElementById("viewTitle");

  if (!user) {
    dashboard.classList.add("hidden");
    title.textContent = "Crew Calendar";
    return;
  }

  dashboard.classList.remove("hidden");
  title.textContent = isAdmin(user) ? "Admin Operations Board" : user.role === "viewer" ? "Calendar View" : "My Crew Schedule";
  ensureActiveSidebarTab(user);

  dashboard.innerHTML = `<div class="single-view" id="singleView"></div>`;

  const singleView = document.getElementById("singleView");
  const visibleShows = visibleShowsForUser(user);
  const shows = filterShowsBySelectedCrew(visibleShows);

  if (state.ui.activeSidebarTab === "calendarPanel") {
    singleView.innerHTML = `<section class="panel" id="calendarPanel"></section>`;
    renderCalendar(user, shows);
    return;
  }

  if (state.ui.activeSidebarTab === "showFormPanel" && isAdmin(user)) {
    singleView.innerHTML = `<section class="panel" id="showFormPanel"></section>`;
    renderShowForm();
    return;
  }

  if (state.ui.activeSidebarTab === "showsPanel") {
    singleView.innerHTML = `<section class="panel" id="showsPanel"></section>`;
    renderShowsList(user, shows, visibleShows);
    return;
  }

  if (state.ui.activeSidebarTab === "crewAdminPanel" && isAdmin(user)) {
    singleView.innerHTML = `<section class="panel" id="crewAdminPanel"></section>`;
    renderCrewAdminPanel();
    return;
  }

  singleView.innerHTML = `<section class="panel" id="legendPanel"></section>`;
  renderLegend(user);
}

function renderCalendarToolbarControls() {
  return `
    <div class="calendar-toolbar-controls">
      <div class="calendar-left-group">
        <div class="view-switch">
          <button class="ghost small calendar-nav-button ${state.view.mode === "month" ? "is-active" : ""}" data-view-mode="month">Month</button>
          <button class="ghost small calendar-nav-button ${state.view.mode === "week" ? "is-active" : ""}" data-view-mode="week">Week</button>
          <button class="ghost small calendar-nav-button ${state.view.mode === "day" ? "is-active" : ""}" data-view-mode="day">Day</button>
        </div>
      </div>
      <div class="calendar-center-group">
        <button class="ghost small calendar-nav-button" data-range-nav="prev">Previous</button>
        <button class="ghost small calendar-nav-button" data-range-nav="today">Today</button>
        <button class="ghost small calendar-nav-button" data-range-nav="next">Next</button>
      </div>
      <div class="calendar-filter-group">
        ${renderCrewFilterControl("calendarCrewFilter", { label: "Crew", allLabel: "All Crew", compact: true, includeUnassigned: true })}
      </div>
    </div>
  `;
}

function wireCalendarToolbarControls(panel) {
  panel.querySelectorAll("[data-view-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      state.view.mode = button.dataset.viewMode;
      saveState(state);
      renderDashboard();
    });
  });

  panel.querySelectorAll("[data-range-nav]").forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.rangeNav;
      if (action === "today") {
        syncViewDateParts(new Date());
      } else {
        const focusDate = getFocusDate();
        const delta = action === "prev" ? -1 : 1;
        const nextFocus = state.view.mode === "month"
          ? new Date(focusDate.getFullYear(), focusDate.getMonth() + delta, 1)
          : addDays(focusDate, state.view.mode === "week" ? delta * 7 : delta);
        syncViewDateParts(nextFocus);
      }

      saveState(state);
      renderDashboard();
    });
  });

  panel.querySelector("#calendarCrewFilter")?.addEventListener("change", (event) => {
    state.ui.selectedCrewFilter = event.currentTarget.value;
    saveState(state);
    renderDashboard();
  });
}

function renderCalendar(user, shows) {
  if (state.view.mode === "week") {
    renderWeekCalendar(user, shows);
    return;
  }

  if (state.view.mode === "day") {
    renderDayCalendar(user, shows);
    return;
  }

  const panel = document.getElementById("calendarPanel");
  const firstDay = new Date(state.view.year, state.view.month, 1);
  const startDay = firstDay.getDay();
  const daysInMonth = new Date(state.view.year, state.view.month + 1, 0).getDate();
  const monthStartKey = dateKey(firstDay);
  const monthEndKey = dateKey(new Date(state.view.year, state.view.month + 1, 0));
  const today = new Date();

  const cells = [];
  for (let i = 0; i < startDay; i += 1) {
    cells.push({ muted: true, label: "", dateKey: null, shows: [] });
  }

  for (let day = 1; day <= daysInMonth; day += 1) {
    const dateKey = `${state.view.year}-${String(state.view.month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const dayShows = getShowsForDate(shows, dateKey);
    cells.push({ muted: false, label: day, dateKey, shows: dayShows });
  }

  while (cells.length % 7 !== 0) {
    cells.push({ muted: true, label: "", dateKey: null, shows: [] });
  }

  const weeks = [];
  for (let index = 0; index < cells.length; index += 7) {
    weeks.push(cells.slice(index, index + 7));
  }
  panel.innerHTML = `
    <div class="calendar-toolbar">
      <div>
        <h3>${monthLabel(state.view.year, state.view.month)}</h3>
        <p class="muted-note">Shows are tinted by assigned crew color. Switch to week or day for time lanes.</p>
      </div>
      ${renderCalendarToolbarControls()}
    </div>
    <div class="calendar-grid">
      ${weekdayLabels.map((day) => `<div class="weekday">${day}</div>`).join("")}
      ${weeks.map((week) => {
        const weekDateKeys = week.map((cell) => cell.dateKey).filter(Boolean);
        const weekStartKey = weekDateKeys[0] || monthStartKey;
        const weekEndKey = weekDateKeys[weekDateKeys.length - 1] || monthEndKey;
        const weekKey = weekStartKey;
        const weekExpanded = Boolean(state.ui.expandedCalendarWeeks[weekKey]);
        const weekShows = shows
          .filter((show) => getCalendarBlockEndDate(show) >= weekStartKey && getCalendarBlockStartDate(show) <= weekEndKey)
          .sort((a, b) => getCalendarBlockStartDate(a).localeCompare(getCalendarBlockStartDate(b)) || getCalendarBlockEndDate(a).localeCompare(getCalendarBlockEndDate(b)) || (a.showName || "").localeCompare(b.showName || ""));
        const lanes = [];
        const hiddenCounts = new Array(7).fill(0);
        const overflowCounts = new Array(7).fill(0);
        const bars = weekShows.map((show) => {
          const clippedStart = getCalendarBlockStartDate(show) > weekStartKey ? getCalendarBlockStartDate(show) : weekStartKey;
          const clippedEnd = getCalendarBlockEndDate(show) < weekEndKey ? getCalendarBlockEndDate(show) : weekEndKey;
          const startIndex = week.findIndex((cell) => cell.dateKey === clippedStart);
          const endIndex = week.findIndex((cell) => cell.dateKey === clippedEnd);
          if (startIndex === -1 || endIndex === -1) return null;
          let laneIndex = lanes.findIndex((laneEnd) => laneEnd < startIndex);
          if (laneIndex === -1) {
            laneIndex = lanes.length;
            lanes.push(endIndex);
          } else {
            lanes[laneIndex] = endIndex;
          }
          return {
            show,
            startIndex,
            endIndex,
            laneIndex,
            clippedStart: clippedStart !== getCalendarBlockStartDate(show),
            clippedEnd: clippedEnd !== getCalendarBlockEndDate(show)
          };
        }).filter(Boolean);
        const collapsedVisibleLanes = 3;
        bars.forEach((bar) => {
          if (bar.laneIndex >= collapsedVisibleLanes) {
            for (let index = bar.startIndex; index <= bar.endIndex; index += 1) {
              overflowCounts[index] += 1;
            }
          }
        });
        const visibleLanes = weekExpanded ? Math.max(lanes.length, 1) : Math.min(Math.max(lanes.length, 1), collapsedVisibleLanes);
        const visibleBars = bars.filter((bar) => {
          if (bar.laneIndex < visibleLanes) return true;
          for (let index = bar.startIndex; index <= bar.endIndex; index += 1) {
            hiddenCounts[index] += 1;
          }
          return false;
        });
        const laneSpace = (visibleLanes * 16) + 8;
        const barMarkup = visibleBars.map((bar) => renderMonthRangeBar(bar.show, user, bar.startIndex, bar.endIndex, bar.laneIndex, bar.clippedStart, bar.clippedEnd)).join("");

        return `
          <div class="month-week">
            <div class="week-days" style="--lane-space:${laneSpace}px;">
              ${barMarkup}
              ${week.map((cell) => {
                const isToday = cell.dateKey && new Date(`${cell.dateKey}T00:00:00`).toDateString() === today.toDateString();
                const cellIndex = week.indexOf(cell);
                const hiddenCount = cell.dateKey ? hiddenCounts[cellIndex] : 0;
                const overflowCount = cell.dateKey ? overflowCounts[cellIndex] : 0;
                return `
                  <article class="calendar-day ${cell.muted ? "muted" : ""} ${isToday ? "today" : ""}" ${cell.dateKey ? `data-date-key="${cell.dateKey}"` : ""}>
                    <span class="day-number">${cell.label || ""}</span>
                    ${cell.dateKey && overflowCount > 0 ? `<button type="button" class="calendar-more-toggle" data-week-key="${weekKey}" data-state="${weekExpanded ? "less" : "more"}">${weekExpanded ? "Less.." : `${hiddenCount || overflowCount} more..`}</button>` : ""}
                  </article>
                `;
              }).join("")}
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;

  if (isAdmin(user)) {
    wireCalendarDragAndDrop(panel);
  }
  panel.querySelectorAll("[data-calendar-show-id]").forEach((node) => {
    node.addEventListener("click", () => {
      const show = shows.find((item) => item.id === node.dataset.calendarShowId);
      if (!show) return;
      ensureUiState();
      state.ui.selectedCalendarShowId = show.id;
      state.ui.calendarReturnMode = "month";
      state.view.mode = "day";
      syncViewDateParts(parseDateKey(node.dataset.focusDate || getShowStartDate(show)));
      saveState(state);
      renderDashboard();
    });
  });
  panel.querySelectorAll("[data-week-key]").forEach((button) => {
    button.addEventListener("click", () => {
      ensureUiState();
      const key = button.dataset.weekKey;
      state.ui.expandedCalendarWeeks[key] = !state.ui.expandedCalendarWeeks[key];
      saveState(state);
      renderDashboard();
    });
  });
  wireCalendarToolbarControls(panel);
}

function getShowDisplayMeta(show, user) {
  const crewUsers = show.assignments
    .map((assignment) => getUserById(assignment.crewId))
    .filter(Boolean);
  const palette = crewUsers.map((crewUser) => crewUser.color).filter(Boolean);
  const color = palette.length > 1
    ? `linear-gradient(135deg, ${palette.join(", ")})`
    : (palette[0] || "linear-gradient(135deg, #575e70, #8b93a5)");
  const visibleAssignees = user.role === "crew"
    ? crewUsers.filter((crewUser) => crewUser.id === user.id)
    : crewUsers;
  return { crewUsers, visibleAssignees, color };
}

function renderMonthRangeBar(show, user, startIndex, endIndex, laneIndex, clippedStart, clippedEnd) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);
  const span = (endIndex - startIndex) + 1;
  const left = `calc(${startIndex} * (100% / 7) + 6px)`;
  const width = `calc(${span} * (100% / 7) - 12px)`;
  const startClass = clippedStart ? "continued-start" : "";
  const endClass = clippedEnd ? "continued-end" : "";
  const locationLabel = show.location ? ` - ${show.location}` : "";
  const crewLabel = visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned";

  return `
    <div class="month-range-bar event-chip ${startClass} ${endClass} ${show.showStatus === "tentative" ? "tentative" : ""} ${isAdmin(user) ? "draggable" : ""}" style="background:${color}; left:${left}; width:${width}; --lane-index:${laneIndex};" data-calendar-show-id="${show.id}" data-focus-date="${getShowStartDate(show)}" ${isAdmin(user) ? `draggable="true" data-show-id="${show.id}" title="Drag to reschedule"` : ""}>
      <p><strong>${show.showName}</strong><span>${locationLabel} (${crewLabel})</span></p>
      <small>${[show.location, visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned"].filter(Boolean).join(" · ")}</small>
    </div>
  `;
}

function renderEventChip(show, user) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);

  return `
    <div class="event-chip ${show.showStatus === "tentative" ? "tentative" : ""} ${isAdmin(user) ? "draggable" : ""}" style="background:${color}" ${isAdmin(user) ? `draggable="true" data-show-id="${show.id}" title="Drag to reschedule"` : ""}>
      <p>${show.showName}</p>
      <small>${[show.location, visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned"].filter(Boolean).join(" · ")}</small>
    </div>
  `;
}

function renderWeekCalendar(user, shows) {
  const panel = document.getElementById("calendarPanel");
  const focusDate = getFocusDate();
  const weekStart = startOfWeek(focusDate);
  const days = Array.from({ length: 7 }, (_, index) => addDays(weekStart, index));
  const weekStartKey = dateKey(weekStart);
  const weekEndKey = dateKey(addDays(weekStart, 6));
  const rangedShows = shows
    .filter((show) => getCalendarBlockEndDate(show) >= weekStartKey && getCalendarBlockStartDate(show) <= weekEndKey)
    .filter((show) => getCalendarBlockStartDate(show) !== getCalendarBlockEndDate(show))
    .sort((a, b) => getCalendarBlockStartDate(a).localeCompare(getCalendarBlockStartDate(b)) || getCalendarBlockEndDate(a).localeCompare(getCalendarBlockEndDate(b)) || (a.showName || "").localeCompare(b.showName || ""));
  const rangeLanes = [];
  const rangeMarkup = rangedShows.map((show) => {
    const clippedStart = getCalendarBlockStartDate(show) > weekStartKey ? getCalendarBlockStartDate(show) : weekStartKey;
    const clippedEnd = getCalendarBlockEndDate(show) < weekEndKey ? getCalendarBlockEndDate(show) : weekEndKey;
    const startIndex = days.findIndex((day) => dateKey(day) === clippedStart);
    const endIndex = days.findIndex((day) => dateKey(day) === clippedEnd);
    if (startIndex === -1 || endIndex === -1) return "";
    let laneIndex = rangeLanes.findIndex((laneEnd) => laneEnd < startIndex);
    if (laneIndex === -1) {
      laneIndex = rangeLanes.length;
      rangeLanes.push(endIndex);
    } else {
      rangeLanes[laneIndex] = endIndex;
    }
    return renderWeekRangeBar(show, user, startIndex, endIndex, laneIndex, clippedStart !== getCalendarBlockStartDate(show), clippedEnd !== getCalendarBlockEndDate(show));
  }).join("");
  const rangeHeight = Math.max(rangeLanes.length, 1) * 26;

  panel.innerHTML = `
    <div class="calendar-toolbar">
      <div>
        <h3>${getVisibleRangeLabel()}</h3>
        <p class="muted-note">Week view shows one shared weekly board with broader strips for assigned entries.</p>
      </div>
      ${renderCalendarToolbarControls()}
    </div>
    <section class="week-board">
      <div class="week-board-headers">
        ${days.map((day) => `<div class="lane-day-header ${dateKey(day) === dateKey(new Date()) ? "today" : ""}">${day.toLocaleDateString("en-IN", { weekday: "short" })}<strong>${formatShortDate(day)}</strong></div>`).join("")}
      </div>
      <div class="week-days week-days-board" style="--lane-space:${rangeHeight + 14}px;">
        ${rangeMarkup}
        ${days.map((day) => {
          const dayKey = dateKey(day);
          const singleDayShows = getShowsForDate(shows, dayKey)
            .filter((show) => getCalendarBlockStartDate(show) === getCalendarBlockEndDate(show))
            .sort((a, b) => (a.showName || "").localeCompare(b.showName || ""));
          return `
            <article class="calendar-day week-calendar-day ${dayKey === dateKey(new Date()) ? "today" : ""}" data-date-key="${dayKey}">
              <span class="day-number">${day.getDate()}</span>
              <div class="week-day-show-list">
                ${singleDayShows.map((show) => renderWeekDayChip(show, user)).join("")}
              </div>
            </article>
          `;
        }).join("")}
      </div>
    </section>
  `;

  if (isAdmin(user)) {
    wireCalendarDragAndDrop(panel);
  }
  panel.querySelectorAll("[data-calendar-show-id]").forEach((node) => {
    node.addEventListener("click", () => {
      const show = shows.find((item) => item.id === node.dataset.calendarShowId);
      if (!show) return;
      ensureUiState();
      state.ui.selectedCalendarShowId = show.id;
      state.ui.calendarReturnMode = "week";
      state.view.mode = "day";
      syncViewDateParts(parseDateKey(node.dataset.focusDate || getShowStartDate(show)));
      saveState(state);
      renderDashboard();
    });
  });
  wireCalendarToolbarControls(panel);
}

function renderWeekRangeBar(show, user, startIndex, endIndex, laneIndex, clippedStart, clippedEnd) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);
  const span = (endIndex - startIndex) + 1;
  const left = `calc(${startIndex} * (100% / 7) + 6px)`;
  const width = `calc(${span} * (100% / 7) - 12px)`;
  const startClass = clippedStart ? "continued-start" : "";
  const endClass = clippedEnd ? "continued-end" : "";
  const locationLabel = show.location ? ` - ${show.location}` : "";
  const crewLabel = visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned";

  return `
    <div class="week-range-bar month-range-bar event-chip ${startClass} ${endClass} ${show.showStatus === "tentative" ? "tentative" : ""} ${isAdmin(user) ? "draggable" : ""}" style="background:${color}; left:${left}; width:${width}; --lane-index:${laneIndex};" data-calendar-show-id="${show.id}" data-focus-date="${getShowStartDate(show)}" ${isAdmin(user) ? `draggable="true" data-show-id="${show.id}" title="Drag to reschedule"` : ""}>
      <p><strong>${show.showName}</strong><span>${locationLabel} (${crewLabel})</span></p>
    </div>
  `;
}

function renderWeekDayChip(show, user) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);
  const locationLabel = show.location ? ` - ${show.location}` : "";
  const crewLabel = visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned";
  return `
    <button type="button" class="week-day-chip event-chip ${show.showStatus === "tentative" ? "tentative" : ""} ${isAdmin(user) ? "draggable" : ""}" style="background:${color}" data-calendar-show-id="${show.id}" data-focus-date="${getShowStartDate(show)}" ${isAdmin(user) ? `draggable="true" data-show-id="${show.id}" title="Drag to reschedule"` : ""}>
      <p><strong>${show.showName}</strong><span>${locationLabel} (${crewLabel})</span></p>
      <small>${visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned"}</small>
    </button>
  `;
}

function renderDayCalendar(user, shows) {
  const panel = document.getElementById("calendarPanel");
  const focusDate = getFocusDate();
  const focusKey = dateKey(focusDate);
  const dayShows = getShowsForDate(shows, focusKey)
    .sort((a, b) => getCalendarBlockStartDate(a).localeCompare(getCalendarBlockStartDate(b)) || (a.showName || "").localeCompare(b.showName || ""));
  const selectedCalendarShow = dayShows.find((show) => show.id === state.ui.selectedCalendarShowId) || null;
  const listShows = selectedCalendarShow
    ? dayShows.filter((show) => show.id !== selectedCalendarShow.id)
    : dayShows;

  panel.innerHTML = `
    <div class="calendar-toolbar">
      <div>
        <h3>${getVisibleRangeLabel()}</h3>
        <p class="muted-note">Day view shows all entries for the selected date in a list.</p>
      </div>
      ${renderCalendarToolbarControls()}
    </div>
    <section class="panel day-list-panel">
      <div class="lane-day-header single ${focusKey === dateKey(new Date()) ? "today" : ""}">
        ${focusDate.toLocaleDateString("en-IN", { weekday: "long" })}
        <strong>${formatWeekdayDate(focusDate)}</strong>
      </div>
      ${selectedCalendarShow ? `
        <div class="calendar-show-panel">
          <div class="calendar-show-panel-header">
            <h4>Show Details</h4>
            <button type="button" class="ghost small" id="closeCalendarShowPanel">Close</button>
          </div>
          ${renderShowCard(selectedCalendarShow, user)}
        </div>
      ` : ""}
      <div class="show-list day-show-list">
        ${listShows.length ? listShows.map((show) => renderShowCard(show, user)).join("") : (selectedCalendarShow ? "<p>No other shows scheduled for this date.</p>" : "<p>No shows scheduled for this date.</p>")}
      </div>
    </section>
  `;
  panel.querySelector("#closeCalendarShowPanel")?.addEventListener("click", () => {
    ensureUiState();
    state.ui.selectedCalendarShowId = null;
    state.view.mode = state.ui.calendarReturnMode || "month";
    state.ui.calendarReturnMode = "month";
    saveState(state);
    renderDashboard();
  });
  wireCalendarToolbarControls(panel);
}

function renderTimeLaneDay(day, dayShows, user, showHeader) {
  const totalMinutes = (DAY_END_HOUR - DAY_START_HOUR) * 60;
  const timedShows = [...dayShows].sort((a, b) => parseTimeToMinutes(a.showTime) - parseTimeToMinutes(b.showTime));

  return `
    <div class="lane-column-wrapper">
      ${showHeader ? "" : `<div class="lane-day-header single ${dateKey(day) === dateKey(new Date()) ? "today" : ""}">${day.toLocaleDateString("en-IN", { weekday: "long" })}<strong>${formatWeekdayDate(day)}</strong></div>`}
      <div class="lane-column" data-date-key="${dateKey(day)}">
        ${Array.from({ length: DAY_END_HOUR - DAY_START_HOUR }, (_, hourIndex) => `<div class="lane-hour"></div>`).join("")}
        ${timedShows.map((show) => renderLaneEvent(show, user, totalMinutes)).join("")}
      </div>
    </div>
  `;
}

function renderLaneEvent(show, user, totalMinutes) {
  const crewUsers = show.assignments.map((assignment) => getUserById(assignment.crewId)).filter(Boolean);
  const palette = crewUsers.map((crewUser) => crewUser.color).filter(Boolean);
  const color = palette.length > 1 ? `linear-gradient(135deg, ${palette.join(", ")})` : (palette[0] || "#264653");
  const visibleAssignees = user.role === "crew" ? crewUsers.filter((crewUser) => crewUser.id === user.id) : crewUsers;
  const startMinutes = parseTimeToMinutes(show.showTime);
  const minutesFromStart = Math.max(0, startMinutes - (DAY_START_HOUR * 60));
  const top = (minutesFromStart / totalMinutes) * 100;
  const height = Math.max((SHOW_BLOCK_MINUTES / totalMinutes) * 100, 8);

  return `
    <article class="lane-event event-chip ${isAdmin(user) ? "draggable" : ""}" style="background:${color}; top:${top}%; height:${height}%;" ${isAdmin(user) ? `draggable="true" data-show-id="${show.id}" title="Drag to reschedule"` : ""}>
      <p>${show.showName}</p>
      <small>${show.showTime ? `${show.showTime} · ` : ""}${show.location || "Location TBD"}</small>
      <small>${visibleAssignees.map((crewUser) => crewUser.name).join(", ") || "Unassigned"}</small>
    </article>
  `;
}

function isShowInCurrentRange(show) {
  const showStart = parseDateKey(getCalendarBlockStartDate(show));
  const showEnd = parseDateKey(getCalendarBlockEndDate(show));
  const focusDate = getFocusDate();
  if (state.view.mode === "day") {
    return isDateWithinCalendarBlock(show, dateKey(focusDate));
  }

  if (state.view.mode === "week") {
    const start = startOfWeek(focusDate);
    const end = addDays(start, 6);
    return showEnd >= start && showStart <= end;
  }

  const monthStart = new Date(state.view.year, state.view.month, 1);
  const monthEnd = new Date(state.view.year, state.view.month + 1, 0);
  return showEnd >= monthStart && showStart <= monthEnd;
}

function wireCalendarDragAndDrop(panel) {
  const chips = panel.querySelectorAll("[data-show-id]");
  const dayCells = panel.querySelectorAll("[data-date-key]");

  chips.forEach((chip) => {
    chip.addEventListener("dragstart", (event) => {
      event.dataTransfer.setData("text/plain", chip.dataset.showId);
      event.dataTransfer.effectAllowed = "move";
      chip.classList.add("dragging");
    });

    chip.addEventListener("dragend", () => {
      chip.classList.remove("dragging");
      dayCells.forEach((cell) => cell.classList.remove("drop-target"));
    });
  });

  dayCells.forEach((cell) => {
    cell.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      cell.classList.add("drop-target");
    });

    cell.addEventListener("dragleave", (event) => {
      if (!cell.contains(event.relatedTarget)) {
        cell.classList.remove("drop-target");
      }
    });

    cell.addEventListener("drop", async (event) => {
      event.preventDefault();
      cell.classList.remove("drop-target");
      const showId = event.dataTransfer.getData("text/plain");
      const show = state.shows.find((item) => item.id === showId);
      const nextDate = cell.dataset.dateKey;
      if (!show || !nextDate || getShowStartDate(show) === nextDate) return;

      const endDate = getShowEndDate(show);
      const durationDays = Math.max(0, Math.round((parseDateKey(endDate) - parseDateKey(getShowStartDate(show))) / (1000 * 60 * 60 * 24)));
      show.showDateFrom = nextDate;
      show.showDateTo = dateKey(addDays(parseDateKey(nextDate), durationDays));
      show.showDate = nextDate;
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });
}

function renderLegend(user) {
  const panel = document.getElementById("legendPanel");
  const legendUsers = user.role === "crew" ? [user] : getCrewUsers();

  panel.innerHTML = `
    <div class="stack">
      <div>
        <h3>Crew Colors</h3>
        ${user.role === "crew" ? "" : '<p class="muted-note">Viewer accounts do not appear in crew assignment lists.</p>'}
      </div>
      <div class="legend">
        ${legendUsers.map((crewUser) => `
          <div class="legend-item">
            <div><span class="legend-swatch" style="background:${crewUser.color}"></span><strong>${crewUser.name}</strong></div>
            <div class="meta">${crewUser.phone}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderShowsList(user, shows, sourceShows = shows) {
  const panel = document.getElementById("showsPanel");
  const groupedShows = shows.reduce((groups, show) => {
    const key = getShowStartDate(show).slice(0, 7);
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(show);
    return groups;
  }, {});
  const sourceGroupedShows = sourceShows.reduce((groups, show) => {
    const key = getShowStartDate(show).slice(0, 7);
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(show);
    return groups;
  }, {});

  const monthKeys = Object.keys(sourceGroupedShows).sort();
  const baseYearOptions = [...new Set(monthKeys.map((monthKey) => monthKey.slice(0, 4)))];
  const yearOptions = ["all", ...baseYearOptions];

  if (!yearOptions.includes(state.ui.selectedShowYear)) {
    state.ui.selectedShowYear = "all";
  }

  const filteredMonthKeys = state.ui.selectedShowYear === "all"
    ? monthKeys
    : monthKeys.filter((monthKey) => monthKey.startsWith(state.ui.selectedShowYear));
  const monthOptions = [
    { value: "all", label: "All Months" },
    ...filteredMonthKeys.map((monthKey) => ({
      value: monthKey,
      label: monthGroupLabel(`${monthKey}-01`).split(" ")[0]
    }))
  ];

  if (state.ui.activeShowMonth !== "all" && !filteredMonthKeys.includes(state.ui.activeShowMonth)) {
    state.ui.activeShowMonth = "all";
  }

  const filteredShows = shows.filter((show) => {
    const showMonthKey = getShowStartDate(show).slice(0, 7);
    const yearMatch = state.ui.selectedShowYear === "all" || showMonthKey.startsWith(state.ui.selectedShowYear);
    const monthMatch = state.ui.activeShowMonth === "all" || showMonthKey === state.ui.activeShowMonth;
    return yearMatch && monthMatch;
  });
  const activeMonthShows = sortShows(filteredShows, state.ui.showSortMode);
  const activeMonthLabel = state.ui.activeShowMonth !== "all"
    ? monthGroupLabel(`${state.ui.activeShowMonth}-01`)
    : state.ui.selectedShowYear !== "all"
      ? `All Months in ${state.ui.selectedShowYear}`
      : "All Months";

  panel.innerHTML = `
    <div class="stack">
      <div>
        <h3>${isAdmin(user) ? "All Show Entries" : "Visible Show Entries"}</h3>
      </div>
      ${sourceShows.length ? `
        <div class="shows-toolbar">
          <div class="shows-toolbar-top">
            ${renderCrewFilterControl("showsCrewFilter", { includeUnassigned: true })}
            <label class="sort-control">
              <span>Year</span>
              <select id="showYearFilter">
                ${yearOptions.map((year) => `<option value="${year}" ${state.ui.selectedShowYear === year ? "selected" : ""}>${year === "all" ? "All Years" : year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="showMonthFilter">
                ${monthOptions.map((option) => `<option value="${option.value}" ${state.ui.activeShowMonth === option.value ? "selected" : ""}>${option.label}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Sort By</span>
              <select id="showSortMode">
                <option value="date" ${state.ui.showSortMode === "date" ? "selected" : ""}>Date</option>
                <option value="crew" ${state.ui.showSortMode === "crew" ? "selected" : ""}>Crew</option>
                <option value="status" ${state.ui.showSortMode === "status" ? "selected" : ""}>Show Status</option>
                <option value="city" ${state.ui.showSortMode === "city" ? "selected" : ""}>City</option>
                <option value="client" ${state.ui.showSortMode === "client" ? "selected" : ""}>Client</option>
              </select>
            </label>
            ${isAdmin(user) ? '<button type="button" class="secondary" id="exportMonthButton">Export Excel</button>' : ""}
          </div>
        </div>
        <section class="month-group">
          <header class="month-group-header">
            <h4>${activeMonthLabel}</h4>
            <span class="pill">${activeMonthShows.length} ${activeMonthShows.length === 1 ? "show" : "shows"}</span>
          </header>
          <div class="show-list">
            ${activeMonthShows.length ? activeMonthShows.map((show) => renderShowCard(show, user)).join("") : "<p>No shows available for the selected crew in this month.</p>"}
          </div>
        </section>
      ` : "<p>No shows available in your current view.</p>"}
    </div>
  `;

  const yearFilterSelect = document.getElementById("showYearFilter");
  if (yearFilterSelect) {
    yearFilterSelect.addEventListener("change", () => {
      state.ui.selectedShowYear = yearFilterSelect.value;
      state.ui.activeShowMonth = "all";
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const monthFilterSelect = document.getElementById("showMonthFilter");
  if (monthFilterSelect) {
    monthFilterSelect.addEventListener("change", () => {
      state.ui.activeShowMonth = monthFilterSelect.value;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const sortSelect = document.getElementById("showSortMode");
  if (sortSelect) {
    sortSelect.addEventListener("change", () => {
      state.ui.showSortMode = sortSelect.value;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const crewFilterSelect = document.getElementById("showsCrewFilter");
  if (crewFilterSelect) {
    crewFilterSelect.addEventListener("change", () => {
      state.ui.selectedCrewFilter = crewFilterSelect.value;
      saveState(state);
      renderDashboard();
    });
  }

  const exportButton = document.getElementById("exportMonthButton");
  if (exportButton && isAdmin(user)) {
    exportButton.addEventListener("click", () => {
      const exportKey = state.ui.activeShowMonth !== "all"
        ? state.ui.activeShowMonth
        : state.ui.selectedShowYear !== "all"
          ? state.ui.selectedShowYear
          : "all-shows";
      exportShowsMonthExcel(exportKey, activeMonthShows);
      showToast(`Exported ${activeMonthLabel} for Excel.`);
    });
  }

  if (isAdmin(user)) {
    panel.querySelectorAll("[data-edit-show]").forEach((button) => {
      button.addEventListener("click", () => fillShowForm(button.dataset.editShow));
    });
  }
}

function renderShowCard(show, user) {
  const assignments = show.assignments.map((assignment) => {
    const crewUser = getUserById(assignment.crewId);
    if (!crewUser) return "";
    const amount = canSeeOperatorAmount(user, assignment) ? formatCurrency(assignment.operatorAmount) : "Hidden";
    const onwardTravelDate = assignment.onwardTravelDate ? formatDate(assignment.onwardTravelDate) : "-";
    const returnTravelDate = assignment.returnTravelDate ? formatDate(assignment.returnTravelDate) : "-";
    const onwardTravelSector = assignment.onwardTravelSector || "-";
    const returnTravelSector = assignment.returnTravelSector || "-";
    const notes = assignment.notes || "-";
    return `
      <div class="assignment-card">
        <header>
          <div>
            <strong>${crewUser.name}</strong>
          </div>
          <span class="pill" style="background:${crewUser.color}; color:white;">Crew</span>
        </header>
        <details class="assignment-details">
          <summary>
            <span class="more-label">More..</span>
            <span class="less-label">Less..</span>
          </summary>
          <div class="assignment-details-body">
            <div class="meta">Operator Amount: ${amount}</div>
            <div class="meta">Onward Travel Date: ${onwardTravelDate}</div>
            <div class="meta">Return Travel Date: ${returnTravelDate}</div>
            <div class="meta">Onward Travel Sector: ${onwardTravelSector}</div>
            <div class="meta">Return Travel Sector: ${returnTravelSector}</div>
            <div class="meta">Notes: ${notes}</div>
          </div>
        </details>
      </div>
    `;
  }).join("");

  return `
    <article class="show-card">
      <header>
        <div>
          <h4>${show.showName}</h4>
          <div class="meta">${formatDateRange(getShowStartDate(show), getShowEndDate(show))}</div>
        </div>
        ${isAdmin(user) ? `<button class="secondary small" data-edit-show="${show.id}">Edit</button>` : ""}
      </header>
      <div class="show-banner">
        <span class="show-banner-item">${show.location || "Location TBD"}</span>
        <span class="show-banner-item">${show.client || "Client TBD"}</span>
        ${isAdmin(user) ? `<span class="show-banner-item">${formatCurrency(show.amountShow)}</span>` : ""}
        <span class="show-banner-item">${show.showStatus === "tentative" ? "Tentative" : "Confirmed"}</span>
      </div>
      <div class="stack" style="margin-top:12px;">
        <strong>Assigned Crew</strong>
        <div class="assignment-list">${assignments || "<p class='meta'>No crew assigned.</p>"}</div>
      </div>
    </article>
  `;
}

function renderShowForm() {
  const panel = document.getElementById("showFormPanel");
  const crewOptions = getCrewUsers();
  ensureUiState();
  const editingShow = state.ui.editingShowId
    ? state.shows.find((show) => show.id === state.ui.editingShowId)
    : null;
  const isEditing = Boolean(editingShow);

  panel.innerHTML = `
    <div class="stack">
      <div>
        <div class="form-header">
          <div>
            <h3>${isEditing ? `Editing: ${editingShow.showName}` : "Create New Show"}</h3>
            <p class="muted-note">Only admins can edit show amount and operator amounts.</p>
          </div>
          ${isEditing ? '<span class="pill edit-pill">Edit Mode</span>' : ""}
        </div>
      </div>
      <form id="showForm" class="stack tight">
        <input type="hidden" name="showId">
        <div class="form-grid">
          <label class="field"><span>Show Date From</span><input type="date" name="showDateFrom" required></label>
          <label class="field"><span>Show Date To</span><input type="date" name="showDateTo" required></label>
          <label class="field"><span>Show Name</span><input type="text" name="showName" required></label>
          <label class="field"><span>Client</span><input type="text" name="client"></label>
          <label class="field"><span>Location</span><input type="text" name="location"></label>
          <label class="field"><span>Amount of the Show</span><input type="number" name="amountShow" min="0" step="1"></label>
        </div>
        <div class="field">
          <span>Show Status</span>
          <div class="status-toggle" role="radiogroup" aria-label="Show Status">
            <label class="status-option">
              <input type="radio" name="showStatus" value="confirmed" checked>
              <span>Confirmed</span>
            </label>
            <label class="status-option">
              <input type="radio" name="showStatus" value="tentative">
              <span>Tentative</span>
            </label>
          </div>
        </div>
        <div class="stack tight">
          <strong>Assign Crew Members</strong>
          <div id="assignmentEditor" class="stack tight"></div>
          <button type="button" class="secondary" id="addAssignmentRow">Add Crew Assignment</button>
        </div>
        <div class="toolbar">
          <button type="submit">${isEditing ? "Update Show" : "Save Show"}</button>
          <button type="button" class="ghost" id="resetShowForm">${isEditing ? "Cancel Edit" : "Clear"}</button>
          ${isEditing ? '<button type="button" class="danger" id="deleteShowButton">Delete Show</button>' : ""}
        </div>
      </form>
    </div>
  `;

  const assignmentEditor = document.getElementById("assignmentEditor");

  function addAssignmentRow(assignment = null) {
    const row = document.createElement("div");
    row.className = "form-grid";
    row.innerHTML = `
      <label class="field">
        <span>Crew</span>
        <select name="assignmentCrew">
          <option value="">Select crew</option>
          ${crewOptions.map((crewUser) => `<option value="${crewUser.id}" ${assignment?.crewId === crewUser.id ? "selected" : ""}>${crewUser.name}</option>`).join("")}
        </select>
      </label>
      <label class="field">
        <span>Amount of the Operator</span>
        <input type="number" name="assignmentAmount" min="0" step="1" value="${assignment?.operatorAmount ?? ""}">
      </label>
      <label class="field">
        <span>Onward Travel Date</span>
        <input type="date" name="assignmentOnwardTravelDate" value="${assignment?.onwardTravelDate ?? ""}">
      </label>
      <label class="field">
        <span>Return Travel Date</span>
        <input type="date" name="assignmentReturnTravelDate" value="${assignment?.returnTravelDate ?? ""}">
      </label>
      <label class="field">
        <span>Onward Travel Sector</span>
        <input type="text" name="assignmentOnwardTravelSector" value="${assignment?.onwardTravelSector ?? ""}">
      </label>
      <label class="field">
        <span>Return Travel Sector</span>
        <input type="text" name="assignmentReturnTravelSector" value="${assignment?.returnTravelSector ?? ""}">
      </label>
      <label class="field">
        <span>Notes</span>
        <input type="text" name="assignmentNotes" value="${assignment?.notes ?? ""}">
      </label>
      <button type="button" class="ghost small remove-assignment">Remove</button>
    `;
    row.querySelector(".remove-assignment").addEventListener("click", () => row.remove());
    assignmentEditor.append(row);
  }

  addAssignmentRow();

  document.getElementById("addAssignmentRow").addEventListener("click", () => addAssignmentRow());
  document.getElementById("resetShowForm").addEventListener("click", () => {
    resetEditingState();
    saveState(state);
    renderDashboard();
  });

  if (isEditing) {
    document.getElementById("deleteShowButton").addEventListener("click", async () => {
      const confirmDelete = window.confirm(`Delete "${editingShow.showName}"?`);
      if (!confirmDelete) return;
      state.shows = state.shows.filter((show) => show.id !== editingShow.id);
      resetEditingState();
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
        showToast("Show deleted.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  }

  document.getElementById("showForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const formData = new FormData(form);
    const showId = formData.get("showId").toString();
    const assignmentRows = [...assignmentEditor.children];
    const assignments = assignmentRows
      .map((row) => ({
        crewId: row.querySelector('select[name="assignmentCrew"]').value,
        operatorAmount: row.querySelector('input[name="assignmentAmount"]').value,
        onwardTravelDate: row.querySelector('input[name="assignmentOnwardTravelDate"]').value,
        returnTravelDate: row.querySelector('input[name="assignmentReturnTravelDate"]').value,
        onwardTravelSector: row.querySelector('input[name="assignmentOnwardTravelSector"]').value,
        returnTravelSector: row.querySelector('input[name="assignmentReturnTravelSector"]').value,
        notes: row.querySelector('input[name="assignmentNotes"]').value
      }))
      .filter((assignment) => assignment.crewId)
      .map((assignment) => ({
        crewId: assignment.crewId,
        operatorAmount: Number(assignment.operatorAmount || 0),
        onwardTravelDate: assignment.onwardTravelDate,
        returnTravelDate: assignment.returnTravelDate,
        onwardTravelSector: assignment.onwardTravelSector.trim(),
        returnTravelSector: assignment.returnTravelSector.trim(),
        notes: assignment.notes.trim()
      }));

    const uniqueCrewIds = new Set(assignments.map((assignment) => assignment.crewId));
    if (uniqueCrewIds.size !== assignments.length) {
      alert("Each crew member can only be assigned once per show.");
      return;
    }

    const payload = {
      id: showId || uid("show"),
      showDateFrom: formData.get("showDateFrom").toString(),
      showDateTo: formData.get("showDateTo").toString(),
      showDate: formData.get("showDateFrom").toString(),
      showStatus: formData.get("showStatus").toString() === "tentative" ? "tentative" : "confirmed",
      showName: formData.get("showName").toString().trim(),
      client: formData.get("client").toString().trim(),
      location: formData.get("location").toString().trim(),
      venue: editingShow?.venue || "",
      showTime: editingShow?.showTime || "",
      amountShow: Number(formData.get("amountShow") || 0),
      assignments
    };

    if (payload.showDateTo < payload.showDateFrom) {
      alert("Show Date To cannot be earlier than Show Date From.");
      return;
    }

    const existingIndex = state.shows.findIndex((show) => show.id === payload.id);
    if (existingIndex >= 0) {
      state.shows[existingIndex] = payload;
    } else {
      state.shows.push(payload);
    }

    resetEditingState();
    try {
      await syncAdminState();
      saveState(state);
      renderDashboard();
      showToast(existingIndex >= 0 ? "Show updated." : "Show created.");
    } catch (error) {
      showToast(error.message);
      await refreshFromServer();
      renderDashboard();
    }
  });
}

function fillShowForm(showId) {
  const show = state.shows.find((item) => item.id === showId);
  if (!show) return;
  ensureUiState();
  state.ui.editingShowId = showId;
  state.ui.activeSidebarTab = "showFormPanel";
  saveState(state);
  renderSidebarTabs();
  renderDashboard();
  const form = document.getElementById("showForm");
  if (!form) return;
  const setFieldValue = (name, value) => {
    const field = form.elements.namedItem(name);
    if (field) {
      field.value = value ?? "";
    }
  };

  setFieldValue("showId", show.id);
  setFieldValue("showDateFrom", getShowStartDate(show));
  setFieldValue("showDateTo", getShowEndDate(show));
  setFieldValue("showName", show.showName);
  setFieldValue("client", show.client);
  setFieldValue("location", show.location);
  setFieldValue("amountShow", show.amountShow);
  const statusFields = form.querySelectorAll('input[name="showStatus"]');
  statusFields.forEach((field) => {
    field.checked = field.value === (show.showStatus === "tentative" ? "tentative" : "confirmed");
  });

  const editor = document.getElementById("assignmentEditor");
  editor.innerHTML = "";

  if (!show.assignments.length) {
    document.getElementById("addAssignmentRow").click();
    form.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  show.assignments.forEach((assignment) => {
    const addButton = document.getElementById("addAssignmentRow");
    addButton.click();
    const row = editor.lastElementChild;
    row.querySelector('select[name="assignmentCrew"]').value = assignment.crewId;
    row.querySelector('input[name="assignmentAmount"]').value = assignment.operatorAmount;
    row.querySelector('input[name="assignmentOnwardTravelDate"]').value = assignment.onwardTravelDate || "";
    row.querySelector('input[name="assignmentReturnTravelDate"]').value = assignment.returnTravelDate || "";
    row.querySelector('input[name="assignmentOnwardTravelSector"]').value = assignment.onwardTravelSector || "";
    row.querySelector('input[name="assignmentReturnTravelSector"]').value = assignment.returnTravelSector || "";
    row.querySelector('input[name="assignmentNotes"]').value = assignment.notes || "";
  });

  form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderApprovalsSection() {
  const pending = getPendingUsers();

  return `
    <div class="stack">
      <div>
        <h3>Pending Approvals</h3>
        <p class="muted-note">Admin, crew, and view-only users must be approved before login.</p>
      </div>
      <div class="approval-list">
        ${pending.length ? pending.map((user) => `
          <article class="show-card">
            <header>
              <div>
                <strong>${user.name}</strong>
                <div class="meta">${user.email} · ${user.phone}</div>
              </div>
              <span class="pill">${user.role === "viewer" ? "View Only" : user.role === "admin" ? "Admin" : "Crew"}</span>
            </header>
            ${(user.role === "crew" || user.role === "admin") && user.color ? `<div class="meta">Requested color <span class="legend-swatch" style="background:${user.color}"></span>${user.color}</div>` : ""}
            <div class="toolbar" style="margin-top:12px;">
              <button class="small" data-approve="${user.id}">Approve</button>
              <button class="secondary small" data-reject="${user.id}">Reject</button>
            </div>
          </article>
        `).join("") : "<p>No pending account approvals.</p>"}
      </div>
    </div>
  `;
}

function wireApprovalActions(scope = document) {
  scope.querySelectorAll("[data-approve]").forEach((button) => {
    button.addEventListener("click", async () => {
      const user = getUserById(button.dataset.approve);
      if (!user) return;
      if ((user.role === "crew" || user.role === "admin") && getTakenColors(user.id).includes(user.color)) {
        alert("That color is no longer available. Ask the crew member to register again with another color.");
        return;
      }
      user.approved = true;
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });

  scope.querySelectorAll("[data-reject]").forEach((button) => {
    button.addEventListener("click", async () => {
      state.users = state.users.filter((user) => user.id !== button.dataset.reject);
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });
}

function renderCrewAdminPanel() {
  const panel = document.getElementById("crewAdminPanel");
  const approvedCrew = sortUsersByName(getApprovedCrewOnlyUsers());
  const approvedAdmins = sortUsersByName(getApprovedAdminUsers());
  const approvedViewers = sortUsersByName(state.users.filter((user) => user.role === "viewer" && user.approved));
  const assignableTeam = sortUsersByName(getCrewUsers());

  panel.innerHTML = `
    <div class="stack">
      <div>
        <h3>Crew Management</h3>
        <p class="muted-note">Admins are also assignable in shows. Use this tab to add or remove crew accounts.</p>
      </div>
      <form id="adminCrewCreateForm" class="stack tight">
        <div class="form-grid">
          <label class="field"><span>Full Name</span><input type="text" name="name" required></label>
          <label class="field"><span>Email</span><input type="email" name="email" required></label>
          <label class="field"><span>Phone</span><input type="tel" name="phone" required></label>
          <label class="field"><span>Password</span><input type="password" name="password" minlength="8" required></label>
        </div>
        <label class="field">
          <span>Crew Color</span>
          <div id="adminCrewColorChoices" class="color-grid"></div>
        </label>
        <p class="muted-note">Use 8+ characters with uppercase, lowercase, and a number.</p>
        <div class="toolbar">
          <button type="submit">Add Crew Member</button>
        </div>
        <div id="adminCrewMessage" class="message"></div>
      </form>
      <form id="adminViewerCreateForm" class="stack tight">
        <div class="form-header">
          <div>
            <h3>Add View Only Account</h3>
            <p class="muted-note">View-only users can log in and inspect the calendar, but they never appear in crew assignments.</p>
          </div>
        </div>
        <div class="form-grid">
          <label class="field"><span>Full Name</span><input type="text" name="name" required></label>
          <label class="field"><span>Email</span><input type="email" name="email" required></label>
          <label class="field"><span>Phone</span><input type="tel" name="phone" required></label>
          <label class="field"><span>Password</span><input type="password" name="password" minlength="8" required></label>
        </div>
        <p class="muted-note">Use 8+ characters with uppercase, lowercase, and a number.</p>
        <div class="toolbar">
          <button type="submit">Add View Only User</button>
        </div>
        <div id="adminViewerMessage" class="message"></div>
      </form>
      <div class="stack">
        <div class="form-header">
          <div>
            <h3>Assignable Team</h3>
            <p class="muted-note">Admins and approved crew members appear in show assignments.</p>
          </div>
          <span class="pill">${assignableTeam.length} members</span>
        </div>
        <div class="legend">
          ${assignableTeam.map((member) => `
            <div class="legend-item">
              <div>
                <span class="legend-swatch" style="background:${member.color || "#264653"}"></span>
                <strong>${member.name}</strong>
              </div>
              <div class="meta">${member.role === "admin" ? "Admin" : "Crew"}</div>
            </div>
          `).join("") || "<p>No assignable team members yet.</p>"}
        </div>
      </div>
      <div class="stack">
        <div class="form-header">
          <div>
            <h3>Approved Admin Accounts</h3>
            <p class="muted-note">Admin removal is locked behind confirmation and at least one approved admin must remain.</p>
          </div>
          <span class="pill">${approvedAdmins.length} admins</span>
        </div>
        <div class="approval-list">
          ${approvedAdmins.length ? approvedAdmins.map((member) => `
            <article class="show-card">
              <header>
                <div>
                  <strong>${member.name}</strong>
                  <div class="meta">${member.email} · ${member.phone}</div>
                </div>
                <span class="pill">Admin</span>
              </header>
              <div class="toolbar" style="margin-top:12px;">
                ${member.id === getCurrentUser()?.id
                  ? '<button type="button" class="secondary small" disabled>Current Admin</button>'
                  : '<button type="button" class="danger small" data-remove-admin="' + member.id + '">Remove Admin</button>'}
              </div>
            </article>
          `).join("") : "<p>No approved admin accounts found.</p>"}
        </div>
      </div>
      <div class="stack">
        <div class="form-header">
          <div>
            <h3>Approved Crew Accounts</h3>
            <p class="muted-note">Removing a crew member also removes them from future assignments.</p>
          </div>
          <span class="pill">${approvedCrew.length} crew</span>
        </div>
        <div class="approval-list">
          ${approvedCrew.length ? approvedCrew.map((member) => `
            <article class="show-card">
              <header>
                <div>
                  <strong>${member.name}</strong>
                  <div class="meta">${member.email} · ${member.phone}</div>
                </div>
                <span class="pill" style="background:${member.color}; color:white;">Crew</span>
              </header>
              <div class="toolbar" style="margin-top:12px;">
                <button type="button" class="danger small" data-remove-crew="${member.id}">Remove Crew</button>
              </div>
            </article>
          `).join("") : "<p>No approved crew accounts yet.</p>"}
        </div>
      </div>
      <div class="stack">
        <div class="form-header">
          <div>
            <h3>Approved View Only Accounts</h3>
            <p class="muted-note">View-only users can inspect schedules but are excluded from crew assignment lists.</p>
          </div>
          <span class="pill">${approvedViewers.length} viewers</span>
        </div>
        <div class="approval-list">
          ${approvedViewers.length ? approvedViewers.map((member) => `
            <article class="show-card">
              <header>
                <div>
                  <strong>${member.name}</strong>
                  <div class="meta">${member.email} · ${member.phone}</div>
                </div>
                <span class="pill">View Only</span>
              </header>
              <div class="toolbar" style="margin-top:12px;">
                <button type="button" class="danger small" data-remove-viewer="${member.id}">Remove View Only</button>
              </div>
            </article>
          `).join("") : "<p>No approved view-only accounts yet.</p>"}
        </div>
      </div>
      ${renderApprovalsSection()}
    </div>
  `;

  renderColorChoices(null, "adminCrewColorChoices");
  wireCrewAdminPanel();
  wireApprovalActions(panel);
}

function sortUsersByName(users) {
  return [...users].sort((a, b) => a.name.localeCompare(b.name));
}

function wireCrewAdminPanel() {
  const form = document.getElementById("adminCrewCreateForm");
  const message = document.getElementById("adminCrewMessage");
  const viewerForm = document.getElementById("adminViewerCreateForm");
  const viewerMessage = document.getElementById("adminViewerMessage");
  const currentUser = getCurrentUser();

  if (form) {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const data = new FormData(form);
      const email = data.get("email").toString().trim().toLowerCase();
      const selectedColorButton = document.querySelector("#adminCrewColorChoices .color-option.selected");

      if (state.users.some((user) => user.email.toLowerCase() === email)) {
        message.textContent = "That email already exists.";
        return;
      }

      if (!selectedColorButton) {
        message.textContent = "Choose an available crew color.";
        return;
      }

      try {
        const payload = await apiRequest("/api/admin/add-crew", {
          method: "POST",
          body: JSON.stringify({
            name: data.get("name").toString().trim(),
            email,
            phone: data.get("phone").toString().trim(),
            password: data.get("password").toString(),
            color: selectedColorButton.dataset.color
          })
        });
        applyServerState(payload);
        saveState(state);
        renderDashboard();
        showToast("Crew member added.");
      } catch (error) {
        message.textContent = error.message;
      }
    });
  }

  if (viewerForm) {
    viewerForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const data = new FormData(viewerForm);
      const email = data.get("email").toString().trim().toLowerCase();

      if (state.users.some((user) => user.email.toLowerCase() === email)) {
        viewerMessage.textContent = "That email already exists.";
        return;
      }

      try {
        const payload = await apiRequest("/api/admin/add-viewer", {
          method: "POST",
          body: JSON.stringify({
            name: data.get("name").toString().trim(),
            email,
            phone: data.get("phone").toString().trim(),
            password: data.get("password").toString()
          })
        });
        applyServerState(payload);
        saveState(state);
        renderDashboard();
        showToast("View-only user added.");
      } catch (error) {
        viewerMessage.textContent = error.message;
      }
    });
  }

  document.querySelectorAll("[data-remove-crew]").forEach((button) => {
    button.addEventListener("click", async () => {
      const crewId = button.dataset.removeCrew;
      if (!crewId || crewId === currentUser?.id) return;
      const crewUser = getUserById(crewId);
      if (!crewUser) return;
      const confirmed = window.confirm(`Remove crew member "${crewUser.name}"?`);
      if (!confirmed) return;

      state.users = state.users.filter((user) => user.id !== crewId);
      state.shows = state.shows.map((show) => ({
        ...show,
        assignments: show.assignments.filter((assignment) => assignment.crewId !== crewId)
      }));
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
        showToast("Crew member removed.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });

  document.querySelectorAll("[data-remove-admin]").forEach((button) => {
    button.addEventListener("click", async () => {
      const adminId = button.dataset.removeAdmin;
      if (!adminId || adminId === currentUser?.id) return;
      const adminUser = getUserById(adminId);
      if (!adminUser) return;

      const approvedAdmins = getApprovedAdminUsers();
      if (approvedAdmins.length <= 1) {
        alert("At least one approved admin must remain.");
        return;
      }

      const confirmation = window.prompt(`Type REMOVE ADMIN to remove ${adminUser.name}.`);
      if (confirmation !== "REMOVE ADMIN") {
        return;
      }

      state.users = state.users.filter((user) => user.id !== adminId);
      state.shows = state.shows.map((show) => ({
        ...show,
        assignments: show.assignments.filter((assignment) => assignment.crewId !== adminId)
      }));

      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
        showToast("Admin account removed.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });

  document.querySelectorAll("[data-remove-viewer]").forEach((button) => {
    button.addEventListener("click", async () => {
      const viewerId = button.dataset.removeViewer;
      const viewerUser = getUserById(viewerId);
      if (!viewerId || !viewerUser) return;
      const confirmed = window.confirm(`Remove view-only user "${viewerUser.name}"?`);
      if (!confirmed) return;

      state.users = state.users.filter((user) => user.id !== viewerId);
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
        showToast("View-only user removed.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    });
  });
}

function render() {
  ensureUiState();
  ensureViewState();
  normalizeState();
  const user = getCurrentUser();
  document.querySelector(".app-shell")?.classList.toggle("logged-in", Boolean(user));
  document.querySelector(".hero-banner")?.classList.toggle("hidden", !user);
  document.querySelector(".topbar")?.classList.toggle("hidden", !user);
  renderSidebarTabs();
  renderAuthPanel();
  renderSessionActions(user);
  renderDashboard();
}

async function initApp() {
  try {
    await refreshFromServer();
  } catch (error) {
    showToast("Server connection failed.");
  }
  render();
}

initApp();
