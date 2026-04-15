const STORAGE_KEY = "pixelbug-calendar-ui-v1";
const COLOR_OPTIONS = [
  "#d93025",
  "#f29900",
  "#f6bf26",
  "#7cb342",
  "#33b679",
  "#4285f4",
  "#7986cb",
  "#8e24aa",
  "#e67c73",
  "#0b8043",
  "#039be5",
  "#c26401",
  "#d81b60"
];
const COLOR_REMAP = {
  "#ee8f8f": "#d93025",
  "#f2b779": "#f29900",
  "#e6cf72": "#f6bf26",
  "#9fd67b": "#7cb342",
  "#73cfc0": "#33b679",
  "#79afea": "#4285f4",
  "#9b92ef": "#7986cb",
  "#c78eeb": "#8e24aa",
  "#ee97bf": "#e67c73",
  "#8fc3a0": "#0b8043",
  "#d9b27c": "#c26401",
  "#7fd6f5": "#039be5",
  "#f59fc1": "#d81b60"
};

const weekdayLabels = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const SHOW_BLOCK_MINUTES = 120;
const DAY_START_HOUR = 6;
const DAY_END_HOUR = 24;
let profileMenuOutsideHandler = null;

function uid(prefix) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function resolveCrewColor(color) {
  if (!color) return color;
  return COLOR_REMAP[color] || color;
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
    .map((assignment) => getAssignmentCrewName(assignment))
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

function safeExcelSheetName(value) {
  return String(value || "Sheet1").replace(/[\\/?*\[\]:]/g, "-").slice(0, 31) || "Sheet1";
}

function safeFileNamePart(value) {
  return String(value || "export").replace(/[^a-z0-9-]+/gi, "-").replace(/^-+|-+$/g, "").toLowerCase() || "export";
}

const INVOICE_EXPORT_COLUMNS = [
  { key: "serial", label: "#", width: 5, type: "number", group: "PixelBug" },
  { key: "showDate", label: "Show Date", width: 13, type: "text", group: "PixelBug" },
  { key: "invoiceNumber", label: "Invoice #", width: 18, type: "text", group: "PixelBug" },
  { key: "invoiceDate", label: "Invoice Date", width: 14, type: "text", group: "PixelBug" },
  { key: "clientName", label: "Client", width: 24, type: "text", group: "PixelBug" },
  { key: "showName", label: "Show Name", width: 26, type: "text", group: "PixelBug" },
  { key: "cityCountry", label: "City/Country", width: 18, type: "text", group: "PixelBug" },
  { key: "showCount", label: "# of Shows", width: 12, type: "number", group: "PixelBug" },
  { key: "basicAmount", label: "Basic Amount", width: 14, type: "number", group: "PixelBug" },
  { key: "gstAmount", label: "GST Amount", width: 14, type: "number", group: "PixelBug" },
  { key: "invoiceAmount", label: "Invoice Amount", width: 15, type: "number", group: "PixelBug" },
  { key: "tdsAmount", label: "TDS Amount", width: 14, type: "number", group: "PixelBug" },
  { key: "receivableAmount", label: "Receivable Amount", width: 17, type: "number", group: "PixelBug" },
  { key: "pixelbugShare", label: "Pixel Bug 10%", width: 15, type: "number", group: "PixelBug" },
  { key: "netAmount", label: "Net Amount", width: 14, type: "number", group: "PixelBug" },
  { key: "operatorRahul", label: "Rahul", width: 13, type: "number", group: "Operator" },
  { key: "operatorAmey", label: "Amey", width: 13, type: "number", group: "Operator" },
  { key: "operatorSachin", label: "Sachin", width: 13, type: "number", group: "Operator" },
  { key: "tpName", label: "T.P. Name", width: 18, type: "text", group: "Operator" },
  { key: "tpAmount", label: "T.P. Amount", width: 14, type: "text", group: "Operator" },
  { key: "tpTds", label: "T.P. TDS", width: 13, type: "text", group: "Operator" },
  { key: "tpPayable", label: "T.P Payable", width: 14, type: "text", group: "Operator" },
  { key: "grossRahul", label: "Rahul", width: 13, type: "number", group: "Gross Amount" },
  { key: "grossAmey", label: "Amey", width: 13, type: "number", group: "Gross Amount" },
  { key: "grossSachin", label: "Sachin", width: 13, type: "number", group: "Gross Amount" }
];

const DEFAULT_INVOICE_EXPORT_COLUMN_KEYS = INVOICE_EXPORT_COLUMNS.map((column) => column.key);
const INVOICE_EXPORT_CORE_NAMES = ["Rahul", "Amey", "Sachin"];

function getInvoiceExportColumns() {
  const selectedKeys = Array.isArray(state.ui.invoiceExportColumns)
    ? state.ui.invoiceExportColumns
    : DEFAULT_INVOICE_EXPORT_COLUMN_KEYS;
  const selectedSet = new Set(selectedKeys);
  const columns = INVOICE_EXPORT_COLUMNS.filter((column) => selectedSet.has(column.key));
  return columns.length ? columns : INVOICE_EXPORT_COLUMNS.filter((column) => DEFAULT_INVOICE_EXPORT_COLUMN_KEYS.includes(column.key));
}

function getExportCoreName(name = "") {
  const normalized = String(name || "").trim().toLowerCase();
  if (!normalized) return "";
  return INVOICE_EXPORT_CORE_NAMES.find((coreName) => {
    const corePattern = new RegExp(`(^|\\s)${coreName.toLowerCase()}(\\s|$)`, "i");
    return corePattern.test(normalized);
  }) || "";
}

function getLinkedInvoiceShows(invoice) {
  return (invoice.lineItems || [])
    .flatMap((item) => parseInvoiceShowIds(item.showId))
    .map((showId) => state.shows.find((show) => show.id === showId))
    .filter(Boolean);
}

function getInvoiceShowCount(invoice, linkedShows = getLinkedInvoiceShows(invoice)) {
  if (linkedShows.length) return linkedShows.length;
  return (invoice.lineItems || []).reduce((sum, item) => sum + Number(item.quantity || 1), 0) || 1;
}

function getInvoiceMonthKey(invoice) {
  return String(invoice.issueDate || invoice.dueDate || "").slice(0, 7) || "undated";
}

function getInvoiceMonthSheetName(monthKey) {
  return monthKey && monthKey !== "undated" ? monthGroupLabel(`${monthKey}-01`) : "Undated";
}

function formatMultilineExportAmounts(values = []) {
  return values
    .map((value) => Number(value || 0))
    .map((value) => value ? Math.round(value * 100) / 100 : "")
    .join("\n");
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
    if (value === "" || value === null || value === undefined) {
      return `<c r="${cellRef}"${styleId ? ` s="${styleId}"` : ""}/>`;
    }
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
      const crewName = getAssignmentCrewName(assignment);
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

      cells.push(makeWorksheetCell(`G${rowNumber}`, crewName));
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

function buildCrewPayoutsSheetXml(rows) {
  const headers = [
    "Show Date",
    "Show Name",
    "Client",
    "Crew",
    "Light Designer",
    "Location",
    "Operator Amount",
    "Status",
    "Notes"
  ];
  const worksheetRows = [];
  worksheetRows.push(`<row r="1">${headers.map((header, index) => makeWorksheetCell(`${excelColumnName(index + 1)}1`, header, "inlineStr", 1)).join("")}</row>`);

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const values = [
      row.showDateLabel,
      row.showName,
      row.clientLabel,
      row.crewName,
      row.lightDesigner,
      row.location,
      row.operatorAmount,
      row.statusLabel,
      row.notes
    ];
    worksheetRows.push(`<row r="${rowNumber}">${values.map((value, valueIndex) => {
      const isNumeric = valueIndex === 6;
      const styleId = !isNumeric && String(value || "").includes("\n") ? 2 : null;
      return makeWorksheetCell(`${excelColumnName(valueIndex + 1)}${rowNumber}`, value, isNumeric ? "n" : "inlineStr", styleId);
    }).join("")}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="15" customWidth="1"/>
    <col min="2" max="2" width="24" customWidth="1"/>
    <col min="3" max="3" width="24" customWidth="1"/>
    <col min="4" max="4" width="20" customWidth="1"/>
    <col min="5" max="5" width="20" customWidth="1"/>
    <col min="6" max="6" width="22" customWidth="1"/>
    <col min="7" max="7" width="15" customWidth="1"/>
    <col min="8" max="8" width="14" customWidth="1"/>
    <col min="9" max="9" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    ${worksheetRows.join("")}
  </sheetData>
</worksheet>`;
}

function exportCrewPayoutsExcel(rows, exportKey = "all") {
  if (!rows.length) return;
  downloadSingleSheetWorkbook("Crew Payouts", buildCrewPayoutsSheetXml(rows), `pixelbug-${safeFileNamePart(exportKey)}-crew-payouts.xlsx`);
}

function buildInvoicesSheetXml(invoices, columns = getInvoiceExportColumns()) {
  const exportColumns = columns.length ? columns : getInvoiceExportColumns();
  const rows = [];
  const mergeRefs = [];
  const groupCells = [];
  let currentGroup = "";
  let groupStartIndex = 1;
  exportColumns.forEach((column, index) => {
    const columnIndex = index + 1;
    if (column.group !== currentGroup) {
      if (currentGroup && columnIndex - 1 > groupStartIndex) {
        mergeRefs.push(`${excelColumnName(groupStartIndex)}1:${excelColumnName(columnIndex - 1)}1`);
      }
      currentGroup = column.group || "";
      groupStartIndex = columnIndex;
      groupCells.push(makeWorksheetCell(`${excelColumnName(columnIndex)}1`, currentGroup, "inlineStr", 1));
    } else {
      groupCells.push(makeWorksheetCell(`${excelColumnName(columnIndex)}1`, "", "inlineStr", 1));
    }
  });
  if (currentGroup && exportColumns.length > groupStartIndex) {
    mergeRefs.push(`${excelColumnName(groupStartIndex)}1:${excelColumnName(exportColumns.length)}1`);
  }

  rows.push(`<row r="1">${groupCells.join("")}</row>`);
  rows.push(`<row r="2">${exportColumns.map((column, index) => makeWorksheetCell(`${excelColumnName(index + 1)}2`, column.label, "inlineStr", 1)).join("")}</row>`);

  invoices.forEach((invoice, index) => {
    const rowNumber = index + 3;
    const placeOfSupply = invoice.details?.placeOfSupply || getClientById(invoice.clientId)?.state || "";
    const calculation = getInvoiceCalculationFromValues(invoice.lineItems || [], placeOfSupply);
    const linkedShows = getLinkedInvoiceShows(invoice);
    const linkedAssignments = linkedShows.flatMap((show) => show.assignments || []);
    const lineDescriptions = (invoice.lineItems || []).map((item) => String(item.description || "").trim()).filter(Boolean).join("\n");
    const showName = getInvoiceLinkedShowNames(invoice).join("\n") || lineDescriptions;
    const showDate = [...new Set(linkedShows.map((show) => formatDateRange(getShowStartDate(show), getShowEndDate(show))).filter((dateLabel) => dateLabel && dateLabel !== "-"))].join("\n");
    const showLocations = [...new Set(linkedShows.map((show) => String(show.location || show.venue || "").trim()).filter(Boolean))].join("\n");
    const basicAmount = calculation.taxableAmount;
    const gstAmount = calculation.taxAmount;
    const invoiceAmount = calculation.totalAmount;
    const tdsAmount = Math.round(basicAmount * 0.10 * 100) / 100;
    const receivableAmount = Math.round((invoiceAmount - tdsAmount) * 100) / 100;
    const pixelbugShare = Math.round(basicAmount * 0.10 * 100) / 100;
    const netAmount = Math.round((basicAmount - pixelbugShare) * 100) / 100;
    const operatorTotals = { Rahul: 0, Amey: 0, Sachin: 0 };
    const grossTotals = { Rahul: 0, Amey: 0, Sachin: 0 };
    const lightDesignerCoreNames = [...new Set(linkedShows
      .flatMap((show) => (show.assignments || []).map((assignment) => getExportCoreName(getUserById(assignment.lightDesignerId)?.name)).filter(Boolean)))];
    const primaryCoreName = lightDesignerCoreNames[0] || getExportCoreName(getInvoiceLightDesignerLabel(invoice));
    if (primaryCoreName) {
      operatorTotals[primaryCoreName] = netAmount;
    }
    const thirdPartyRows = linkedAssignments
      .map((assignment) => {
        const crewName = getAssignmentCrewName(assignment);
        const amount = Math.round(Number(assignment.operatorAmount || 0) * 100) / 100;
        const tds = Math.round(amount * 0.10 * 100) / 100;
        return {
          name: crewName,
          amount,
          tds,
          payable: Math.round((amount - tds) * 100) / 100
        };
      })
      .filter((row) => row.name && !getExportCoreName(row.name));
    const tpName = thirdPartyRows.map((row) => row.name).join("\n");
    const tpAmount = formatMultilineExportAmounts(thirdPartyRows.map((row) => row.amount));
    const tpAmountTotal = thirdPartyRows.reduce((sum, row) => sum + row.amount, 0);
    const tpTds = formatMultilineExportAmounts(thirdPartyRows.map((row) => row.tds));
    const tpPayable = formatMultilineExportAmounts(thirdPartyRows.map((row) => row.payable));
    INVOICE_EXPORT_CORE_NAMES.forEach((coreName) => {
      grossTotals[coreName] = Math.max(0, operatorTotals[coreName] - (coreName === primaryCoreName ? tpAmountTotal : 0));
    });
    const values = {
      serial: index + 1,
      showDate: showDate || formatInvoiceDate(invoice.issueDate),
      invoiceNumber: invoice.invoiceNumber,
      invoiceDate: invoice.issueDate,
      clientName: invoice.clientName,
      showName,
      cityCountry: showLocations || placeOfSupply,
      showCount: getInvoiceShowCount(invoice, linkedShows),
      basicAmount,
      gstAmount,
      invoiceAmount,
      tdsAmount,
      receivableAmount,
      pixelbugShare,
      netAmount,
      operatorRahul: operatorTotals.Rahul || "",
      operatorAmey: operatorTotals.Amey || "",
      operatorSachin: operatorTotals.Sachin || "",
      tpName,
      tpAmount,
      tpTds,
      tpPayable,
      grossRahul: grossTotals.Rahul || "",
      grossAmey: grossTotals.Amey || "",
      grossSachin: grossTotals.Sachin || ""
    };
    rows.push(`<row r="${rowNumber}">${exportColumns.map((exportColumn, cellIndex) => {
      const cellColumn = excelColumnName(cellIndex + 1);
      const cellValue = values[exportColumn.key];
      const cellStyle = typeof cellValue === "string" && cellValue.includes("\n") ? 2 : null;
      return makeWorksheetCell(`${cellColumn}${rowNumber}`, cellValue, exportColumn.type === "number" ? "n" : "inlineStr", cellStyle);
    }).join("")}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    ${exportColumns.map((column, index) => {
      const columnIndex = index + 1;
      return `<col min="${columnIndex}" max="${columnIndex}" width="${column.width || 18}" customWidth="1"/>`;
    }).join("")}
  </cols>
  <sheetData>
    ${rows.join("")}
  </sheetData>
  ${mergeRefs.length ? `<mergeCells count="${mergeRefs.length}">${mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>` : ""}
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

function downloadSingleSheetWorkbook(sheetName, sheetXml, fileName) {
  downloadWorkbook([{ name: sheetName, xml: sheetXml }], fileName);
}

function downloadWorkbook(sheets, fileName) {
  const safeSheets = sheets
    .map((sheet, index) => ({
      name: safeExcelSheetName(sheet.name || `Sheet ${index + 1}`),
      xml: sheet.xml
    }))
    .filter((sheet) => sheet.xml);
  if (!safeSheets.length) return;
  const files = [
    {
      name: "[Content_Types].xml",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${safeSheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}
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
    ${safeSheets.map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}
  </sheets>
</workbook>`
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${safeSheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}
  <Relationship Id="rId${safeSheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
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
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`
    },
    ...safeSheets.map((sheet, index) => ({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      data: sheet.xml
    }))
  ];

  const blob = createZip(files);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(link.href);
}

function exportShowsMonthExcel(monthKey, shows) {
  if (!monthKey || !shows.length) return;
  downloadSingleSheetWorkbook(monthKey, buildShowsSheetXml(shows), `pixelbug-${monthKey}-shows.xlsx`);
}

function buildClientsSheetXml(clients) {
  const headers = [
    "Client Name",
    "State",
    "GSTIN",
    "Contact Person",
    "Contact Email",
    "Contact Phone",
    "Billing Address",
    "Notes",
    "Linked Shows",
    "Linked Invoices"
  ];

  const rows = [];
  rows.push(`<row r="1">${headers.map((header, index) => makeWorksheetCell(`${excelColumnName(index + 1)}1`, header, "inlineStr", 1)).join("")}</row>`);

  clients.forEach((client, index) => {
    const rowNumber = index + 2;
    const linkedShows = state.shows.filter((show) => (show.clientId || getClientByName(show.client)?.id || "") === client.id).length;
    const linkedInvoices = state.invoices.filter((invoice) => (invoice.clientId || getClientByName(invoice.clientName)?.id || "") === client.id).length;
    const values = [
      getClientDisplayName(client),
      client.state || "",
      client.gstin || "",
      client.contactName || "",
      client.contactEmail || "",
      client.contactPhone || "",
      client.billingAddress || "",
      client.notes || "",
      linkedShows,
      linkedInvoices
    ];
    rows.push(`<row r="${rowNumber}">${values.map((value, valueIndex) => {
      const isNumber = valueIndex >= 8;
      return makeWorksheetCell(`${excelColumnName(valueIndex + 1)}${rowNumber}`, value, isNumber ? "n" : "inlineStr", isNumber ? null : 2);
    }).join("")}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="28" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="3" width="22" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
    <col min="5" max="5" width="28" customWidth="1"/>
    <col min="6" max="6" width="18" customWidth="1"/>
    <col min="7" max="7" width="42" customWidth="1"/>
    <col min="8" max="8" width="28" customWidth="1"/>
    <col min="9" max="10" width="14" customWidth="1"/>
  </cols>
  <sheetData>
    ${rows.join("")}
  </sheetData>
</worksheet>`;
}

function exportClientsExcel(clients) {
  if (!clients.length) return;
  downloadSingleSheetWorkbook("Clients", buildClientsSheetXml(clients), "pixelbug-clients.xlsx");
}

function getClientInvoiceYearOptions() {
  return [...new Set((state.invoices || []).map((invoice) => String(invoice.issueDate || "").slice(0, 4)).filter(Boolean))].sort((a, b) => b.localeCompare(a));
}

function getClientInvoiceMonthOptions(selectedYear = "all") {
  const monthKeys = (state.invoices || [])
    .filter((invoice) => selectedYear === "all" || String(invoice.issueDate || "").slice(0, 4) === selectedYear)
    .map((invoice) => String(invoice.issueDate || "").slice(0, 7))
    .filter(Boolean);
  return [...new Set(monthKeys)].sort((a, b) => b.localeCompare(a));
}

function getClientExportRows(clientIds = [], selectedYear = "all", selectedMonth = "all") {
  const clientSet = new Set(clientIds.filter(Boolean));
  const invoices = (state.invoices || []).filter((invoice) => {
    const invoiceClientId = invoice.clientId || getClientByName(invoice.clientName)?.id || "";
    if (clientSet.size && !clientSet.has(invoiceClientId)) return false;
    const invoiceYear = String(invoice.issueDate || "").slice(0, 4);
    const invoiceMonth = String(invoice.issueDate || "").slice(0, 7);
    if (selectedYear !== "all" && invoiceYear !== selectedYear) return false;
    if (selectedMonth !== "all" && invoiceMonth !== selectedMonth) return false;
    return true;
  });

  return invoices.map((invoice) => {
    const client = getClientById(invoice.clientId) || getClientByName(invoice.clientName);
    const linkedShows = getLinkedInvoiceShows(invoice);
    const lightDesigner = getInvoiceLightDesignerLabel(invoice) || "-";
    const crewNames = [...new Set(linkedShows.flatMap((show) => (show.assignments || []).map((assignment) => getAssignmentCrewName(assignment)).filter(Boolean)))];
    const crewAmounts = linkedShows.flatMap((show) => (show.assignments || []).map((assignment) => Number(assignment.operatorAmount || 0)).filter((amount) => amount > 0));
    const paymentEntries = Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : [];
    const paymentDates = paymentEntries.map((payment) => formatInvoiceDate(payment.paymentDate)).filter(Boolean).join("\n");
    const paymentAmounts = paymentEntries.map((payment) => Number(payment.amount || 0)).filter((amount) => amount > 0).join("\n");
    const paymentNotes = paymentEntries.map((payment) => String(payment.note || "").trim()).filter(Boolean).join("\n");
    return {
      clientName: getClientDisplayName(client || { name: invoice.clientName || "" }),
      state: client?.state || "",
      gstin: client?.gstin || invoice.details?.clientGstin || "",
      contactName: client?.contactName || "",
      contactEmail: client?.contactEmail || "",
      contactPhone: client?.contactPhone || "",
      billingAddress: client?.billingAddress || "",
      invoiceNumber: invoice.invoiceNumber || "",
      issueDate: formatInvoiceDate(invoice.issueDate),
      dueDate: formatInvoiceDate(invoice.dueDate),
      invoiceStatus: getInvoiceStatusLabel(invoice),
      paymentStatus: getInvoicePaymentBucket(invoice),
      paymentDates,
      paymentAmounts,
      paymentNotes,
      amountPaid: Number(invoice.amountPaid || 0),
      balanceDue: Number(invoice.balanceDue || 0),
      totalAmount: Number(invoice.totalAmount || 0),
      placeOfSupply: invoice.details?.placeOfSupply || client?.state || "",
      lightDesigner,
      showNames: linkedShows.map((show) => show.showName).filter(Boolean).join("\n"),
      showLocations: linkedShows.map((show) => show.location || show.venue || "").filter(Boolean).join("\n"),
      crewNames: crewNames.join("\n"),
      crewAmounts: crewAmounts.join("\n"),
      invoiceNotes: invoice.notes || ""
    };
  });
}

function buildClientInvoiceExportSheetXml(rows) {
  const headers = [
    "Client Name",
    "State",
    "GSTIN",
    "Contact Person",
    "Contact Email",
    "Contact Phone",
    "Billing Address",
    "Invoice No",
    "Invoice Date",
    "Due Date",
    "Invoice Status",
    "Payment Status",
    "Payment Dates",
    "Payment Amounts",
    "Payment Notes",
    "Amount Paid",
    "Balance Due",
    "Invoice Total",
    "Place of Supply",
    "Light Designer",
    "Show Names",
    "Show Locations",
    "Crew Names",
    "Crew Amounts",
    "Invoice Notes"
  ];

  const worksheetRows = [];
  worksheetRows.push(`<row r="1">${headers.map((header, index) => makeWorksheetCell(`${excelColumnName(index + 1)}1`, header, "inlineStr", 1)).join("")}</row>`);

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const values = [
      row.clientName,
      row.state,
      row.gstin,
      row.contactName,
      row.contactEmail,
      row.contactPhone,
      row.billingAddress,
      row.invoiceNumber,
      row.issueDate,
      row.dueDate,
      row.invoiceStatus,
      row.paymentStatus,
      row.paymentDates,
      row.paymentAmounts,
      row.paymentNotes,
      row.amountPaid,
      row.balanceDue,
      row.totalAmount,
      row.placeOfSupply,
      row.lightDesigner,
      row.showNames,
      row.showLocations,
      row.crewNames,
      row.crewAmounts,
      row.invoiceNotes
    ];
    worksheetRows.push(`<row r="${rowNumber}">${values.map((value, valueIndex) => {
      const isNumeric = [15, 16, 17].includes(valueIndex);
      return makeWorksheetCell(`${excelColumnName(valueIndex + 1)}${rowNumber}`, value, isNumeric ? "n" : "inlineStr", isNumeric ? null : 2);
    }).join("")}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="26" customWidth="1"/>
    <col min="2" max="3" width="18" customWidth="1"/>
    <col min="4" max="6" width="20" customWidth="1"/>
    <col min="7" max="7" width="34" customWidth="1"/>
    <col min="8" max="12" width="16" customWidth="1"/>
    <col min="13" max="15" width="18" customWidth="1"/>
    <col min="16" max="18" width="14" customWidth="1"/>
    <col min="19" max="20" width="18" customWidth="1"/>
    <col min="21" max="24" width="24" customWidth="1"/>
    <col min="25" max="25" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    ${worksheetRows.join("")}
  </sheetData>
</worksheet>`;
}

function exportClientInvoiceExcel(exportKey, rows) {
  if (!rows.length) return;
  downloadSingleSheetWorkbook("Client Invoices", buildClientInvoiceExportSheetXml(rows), `pixelbug-${safeFileNamePart(exportKey)}-client-invoices.xlsx`);
}

function getClientLedgerEntries(clientId) {
  if (!clientId) return [];
  const relatedInvoices = (state.invoices || [])
    .filter((invoice) => (invoice.clientId || getClientByName(invoice.clientName)?.id || "") === clientId);
  const entries = relatedInvoices.flatMap((invoice) => {
    const linkedShows = getLinkedInvoiceShows(invoice);
    const invoiceParticulars = linkedShows.length
      ? linkedShows.map((show) => show.showName).filter(Boolean).join(", ")
      : (String(invoice.notes || "").trim() || "Invoice issued");
    const invoiceEntry = {
      date: invoice.issueDate || "",
      type: "Invoice",
      reference: invoice.invoiceNumber || "",
      particulars: invoiceParticulars,
      debit: Number(invoice.totalAmount || 0),
      credit: 0
    };
    const paymentEntries = Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : [];
    const payments = paymentEntries.map((payment) => ({
      date: payment.paymentDate || invoice.issueDate || "",
      type: "Payment",
      reference: invoice.invoiceNumber || "",
      particulars: String(payment.note || "").trim() || "Payment received",
      debit: 0,
      credit: Number(payment.amount || 0)
    }));
    return [invoiceEntry, ...payments];
  });

  const typeOrder = { Invoice: 0, Payment: 1 };
  entries.sort((a, b) => {
    const dateCompare = String(a.date || "").localeCompare(String(b.date || ""));
    if (dateCompare !== 0) return dateCompare;
    const typeCompare = (typeOrder[a.type] ?? 9) - (typeOrder[b.type] ?? 9);
    if (typeCompare !== 0) return typeCompare;
    return String(a.reference || "").localeCompare(String(b.reference || ""));
  });

  let runningBalance = 0;
  return entries.map((entry) => {
    runningBalance += Number(entry.debit || 0) - Number(entry.credit || 0);
    return {
      ...entry,
      balance: runningBalance
    };
  });
}

function buildClientLedgerSheetXml(client, entries) {
  const headers = [
    "Date",
    "Type",
    "Reference",
    "Particulars",
    "Debit",
    "Credit",
    "Balance"
  ];
  const rows = [];
  rows.push(`<row r="1">${headers.map((header, index) => makeWorksheetCell(`${excelColumnName(index + 1)}1`, header, "inlineStr", 1)).join("")}</row>`);
  entries.forEach((entry, index) => {
    const rowNumber = index + 2;
    const values = [
      formatInvoiceDate(entry.date),
      entry.type,
      entry.reference,
      entry.particulars,
      entry.debit,
      entry.credit,
      entry.balance
    ];
    rows.push(`<row r="${rowNumber}">${values.map((value, valueIndex) => {
      const isNumeric = valueIndex >= 4;
      return makeWorksheetCell(`${excelColumnName(valueIndex + 1)}${rowNumber}`, value, isNumeric ? "n" : "inlineStr", isNumeric ? null : 2);
    }).join("")}</row>`);
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="14" customWidth="1"/>
    <col min="2" max="3" width="18" customWidth="1"/>
    <col min="4" max="4" width="42" customWidth="1"/>
    <col min="5" max="7" width="16" customWidth="1"/>
  </cols>
  <sheetData>
    ${rows.join("")}
  </sheetData>
</worksheet>`;
}

function exportClientLedgerExcel(client) {
  if (!client) return;
  const entries = getClientLedgerEntries(client.id);
  if (!entries.length) return;
  downloadSingleSheetWorkbook(
    "Client Ledger",
    buildClientLedgerSheetXml(client, entries),
    `pixelbug-${safeFileNamePart(getClientDisplayName(client))}-ledger.xlsx`
  );
}

function getClientLedgerMarkup(client) {
  const entries = getClientLedgerEntries(client.id);
  const relatedInvoices = (state.invoices || []).filter((invoice) => (invoice.clientId || getClientByName(invoice.clientName)?.id || "") === client.id);
  const totalBilled = relatedInvoices.reduce((sum, invoice) => sum + Number(invoice.totalAmount || 0), 0);
  const totalCollected = relatedInvoices.reduce((sum, invoice) => sum + Number(invoice.amountPaid || 0), 0);
  const totalOutstanding = relatedInvoices.reduce((sum, invoice) => sum + Number(invoice.balanceDue || 0), 0);
  return `
    <div class="client-ledger-print-page">
      <header class="client-ledger-print-header">
        <div>
          <div class="client-ledger-print-title">Client Ledger</div>
          <div class="client-ledger-print-subtitle">${escapeHtml(getClientDisplayName(client))}</div>
        </div>
        <div class="client-ledger-print-meta">
          <div>${client.state ? escapeHtml(client.state) : "State not added"}</div>
          <div>${client.gstin ? `GSTIN: ${escapeHtml(client.gstin)}` : "GSTIN not added"}</div>
        </div>
      </header>
      <section class="client-ledger-print-summary">
        <div><span>Total Billed</span><strong>${escapeHtml(formatCurrency(totalBilled))}</strong></div>
        <div><span>Collected</span><strong>${escapeHtml(formatCurrency(totalCollected))}</strong></div>
        <div><span>Outstanding</span><strong>${escapeHtml(formatCurrency(totalOutstanding))}</strong></div>
      </section>
      <section class="client-ledger-print-details">
        <div><strong>Contact</strong> ${escapeHtml(client.contactName || "No contact")}${client.contactEmail ? ` · ${escapeHtml(client.contactEmail)}` : ""}${client.contactPhone ? ` · ${escapeHtml(client.contactPhone)}` : ""}</div>
        <div><strong>Billing Address</strong> ${escapeHtml(client.billingAddress || "No billing address saved yet.")}</div>
      </section>
      <table class="client-ledger-print-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Type</th>
            <th>Reference</th>
            <th>Particulars</th>
            <th>Debit</th>
            <th>Credit</th>
            <th>Balance</th>
          </tr>
        </thead>
        <tbody>
          ${entries.length ? entries.map((entry) => `
            <tr>
              <td>${escapeHtml(formatInvoiceDate(entry.date))}</td>
              <td>${escapeHtml(entry.type)}</td>
              <td>${escapeHtml(entry.reference || "-")}</td>
              <td>${escapeHtml(entry.particulars || "-")}</td>
              <td>${entry.debit ? escapeHtml(formatCurrency(entry.debit)) : "-"}</td>
              <td>${entry.credit ? escapeHtml(formatCurrency(entry.credit)) : "-"}</td>
              <td>${escapeHtml(formatCurrency(entry.balance))}</td>
            </tr>
          `).join("") : `<tr><td colspan="7">No ledger transactions yet.</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function getClientLedgerPrintStyles() {
  return `
    @page { size: A4; margin: 12mm; }
    * { box-sizing: border-box; }
    body { margin: 0; color: #17212b; font-family: "IBM Plex Sans", sans-serif; background: #fff; }
    .client-ledger-print-page { width: 100%; font-size: 11px; }
    .client-ledger-print-header { display: flex; justify-content: space-between; gap: 16px; align-items: start; border-bottom: 2px solid #d8e0ea; padding-bottom: 10px; margin-bottom: 12px; }
    .client-ledger-print-title { font-family: "Space Grotesk", sans-serif; font-size: 24px; font-weight: 700; }
    .client-ledger-print-subtitle { margin-top: 2px; font-size: 14px; font-weight: 600; }
    .client-ledger-print-meta { text-align: right; color: #667789; font-size: 10px; line-height: 1.4; }
    .client-ledger-print-summary { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; margin-bottom: 12px; }
    .client-ledger-print-summary div { padding: 10px; border: 1px solid #dfe6ef; border-radius: 12px; background: #fbfcfe; }
    .client-ledger-print-summary span { display: block; color: #667789; font-size: 10px; text-transform: uppercase; letter-spacing: .08em; }
    .client-ledger-print-summary strong { display: block; margin-top: 4px; font-size: 15px; }
    .client-ledger-print-details { margin-bottom: 12px; color: #4a5b6d; line-height: 1.5; }
    .client-ledger-print-details div { margin-bottom: 4px; }
    .client-ledger-print-table { width: 100%; border-collapse: collapse; table-layout: auto; }
    .client-ledger-print-table th, .client-ledger-print-table td { padding: 7px 8px; border-bottom: 1px solid #dfe6ef; text-align: left; vertical-align: top; }
    .client-ledger-print-table th { background: #f8fafc; color: #5c6b7a; font-size: 9px; letter-spacing: .08em; text-transform: uppercase; }
  `;
}

function printClientLedger(clientId) {
  const client = getClientById(clientId);
  if (!client) return;
  const ledgerWindow = window.open("", "_blank");
  if (!ledgerWindow) {
    showToast("Unable to open print window.");
    return;
  }
  const printTitle = `${escapeHtml(getClientDisplayName(client))} Ledger`;
  ledgerWindow.document.open();
  ledgerWindow.document.write(`<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <title>${printTitle}</title>
    <style>${getClientLedgerPrintStyles()}</style>
  </head>
  <body>
    ${getClientLedgerMarkup(client)}
    <script>
      window.addEventListener("load", () => {
        window.setTimeout(() => {
          window.focus();
          window.print();
        }, 250);
      });
    <\/script>
  </body>
</html>`);
  ledgerWindow.document.close();
}

function exportInvoicesExcel(exportKey, invoices, columns = getInvoiceExportColumns(), options = {}) {
  if (!exportKey || !invoices.length) return;
  const monthKeys = [...new Set(invoices.map(getInvoiceMonthKey))].sort();
  const fileName = `pixelbug-${safeFileNamePart(exportKey)}-invoices.xlsx`;
  if (options.splitByMonth && monthKeys.length > 1) {
    const sheets = monthKeys.map((monthKey) => {
      const monthInvoices = invoices.filter((invoice) => getInvoiceMonthKey(invoice) === monthKey);
      return {
        name: getInvoiceMonthSheetName(monthKey),
        xml: buildInvoicesSheetXml(monthInvoices, columns)
      };
    });
    downloadWorkbook(sheets, fileName);
    return;
  }
  const sheetName = monthKeys.length === 1 ? getInvoiceMonthSheetName(monthKeys[0]) : exportKey;
  downloadSingleSheetWorkbook(sheetName, buildInvoicesSheetXml(invoices, columns), fileName);
}

function seedState() {
  return {
    users: [],
    shows: [],
    clients: [],
    invoices: [],
    currentUserId: null,
    google: {
      configured: false,
      connected: false,
      calendarId: "",
      lastSyncAt: "",
      lastError: ""
    },
    view: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(),
      mode: "month",
      focusDate: dateKey(new Date())
    },
    ui: {
      editingShowId: null,
      newShowDate: "",
      showDraftTemplate: null,
      showSubtab: "list",
      showTimelineMode: "active",
      editingInvoiceId: null,
      activeSidebarTab: "calendarPanel",
      activeShowMonth: null,
      selectedShowYear: "all",
      showSortMode: "date",
      showSearchQuery: "",
      showsPage: 1,
      showsPageSize: 10,
      selectedCrewFilter: "all",
      googleEntriesView: "needsCompletion",
      googleEntriesPage: 1,
      googleEntriesPageSize: 10,
      invoiceSearchQuery: "",
      invoiceStatusFilter: "all",
      invoicePaymentFilter: "all",
      invoiceClientFilter: "all",
      invoiceExportYear: "all",
      invoiceExportMonth: "all",
      invoiceLightDesignerFilter: "all",
      invoiceExportColumns: DEFAULT_INVOICE_EXPORT_COLUMN_KEYS,
      invoiceSortMode: "issueDate",
      invoiceRegisterPage: 1,
      invoiceRegisterPageSize: 10,
      paymentReconSearchQuery: "",
      paymentReconYear: "all",
      paymentReconMonth: "all",
      paymentReconClient: "all",
      paymentReconPage: 1,
      paymentReconPageSize: 10,
      payoutSearchQuery: "",
      payoutYear: "all",
      payoutMonth: "all",
      payoutCrew: "all",
      payoutClient: "all",
      payoutPage: 1,
      payoutPageSize: 10,
      documentShowsYear: "all",
      documentShowsMonth: "all",
      documentInvoicesYear: "all",
      documentInvoicesMonth: "all",
      documentInvoicesClient: "all",
      documentClientFinancialYear: "all",
      documentClientFinancialMonth: "all",
      documentClientFinancialClient: "all",
      documentLedgerClient: "all",
      documentClientsGstFilter: "all",
      documentPayoutYear: "all",
      documentPayoutMonth: "all",
      documentPayoutCrew: "all",
      documentPayoutClient: "all",
      invoiceDraftShowIds: [],
      invoiceDraftTemplate: null,
      invoiceSubtab: "create",
      markPaymentInvoiceId: null,
      clientsPage: 1,
      clientsPageSize: 10,
      clientsSubtab: "list",
      selectedClientDetailId: null,
      clientSearchQuery: "",
      clientGstFilter: "all",
      clientExportYear: "all",
      clientExportMonth: "all",
      clientExportClientId: "all",
      dirtyForm: null,
      invoicePrintCopies: 1,
      invoiceLineDefaults: {
        description: "",
        sac: ""
      },
      googleArchiveYear: "all",
      googleArchiveMonth: "all",
      expandedCalendarWeeks: {},
      selectedCalendarShowId: null,
      calendarReturnMode: "month",
      themePreference: "system",
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
      ui: parsed.ui || seedState().ui,
      google: seedState().google,
      clients: [],
      invoices: []
    };
  } catch (error) {
    return seedState();
  }
}

function normalizeState() {
  ensureUiState();
  ensureViewState();

  state.clients = Array.isArray(state.clients)
    ? state.clients.map((client) => normalizeClient(client))
    : [];

  state.shows = state.shows.map((show) => {
    const legacyTravelDate = show.travelDate || "";
    const legacyTravelSector = show.travelSector || "";
    const legacyTravelNotes = show.travelNotes || "";

    const assignments = (show.assignments || []).map((assignment) => ({
      ...assignment,
      crewId: assignment.crewId ?? "",
      manualCrewName: assignment.manualCrewName ?? "",
      lightDesignerId: assignment.lightDesignerId ?? "",
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
      clientId: show.clientId || getClientByName(show.client)?.id || "",
      showDateFrom: show.showDateFrom ?? show.showDate ?? "",
      showDateTo: show.showDateTo ?? show.showDateFrom ?? show.showDate ?? "",
      showDate: show.showDateFrom ?? show.showDate ?? "",
      showStatus: show.showStatus === "tentative" ? "tentative" : "confirmed",
      googleEventId: show.googleEventId ?? "",
      googleSyncSource: show.googleSyncSource ?? "",
      googleSyncStatus: show.googleSyncStatus ?? (show.googleEventId ? "synced" : ""),
      googleNotes: show.googleNotes ?? "",
      googleLastSyncedAt: show.googleLastSyncedAt ?? "",
      needsAdminCompletion: Boolean(show.needsAdminCompletion),
      googleArchived: Boolean(show.googleArchived),
      googleArchivedAt: show.googleArchivedAt ?? "",
      googlePinned: Boolean(show.googlePinned),
      assignments,
      travelDate: undefined,
      travelSector: undefined,
      travelNotes: undefined
    };
  });
  state.invoices = Array.isArray(state.invoices)
    ? state.invoices.map((invoice) => ({
        ...invoice,
        clientId: invoice.clientId || getClientByName(invoice.clientName)?.id || "",
        details: normalizeInvoiceDetails(invoice.details),
        lineItems: Array.isArray(invoice.lineItems)
          ? invoice.lineItems.map((item, index) => ({
              ...item,
              showId: item.showId ?? "",
              description: item.description ?? "",
              sac: item.sac ?? "",
              customDetails: item.customDetails ?? "",
              discount: item.discount ?? "",
              discountAmount: Number(item.discountAmount || 0),
              quantity: Number(item.quantity || 0),
              unitRate: Number(item.unitRate || 0),
              amount: Number(item.amount || 0),
              lineOrder: Number.isFinite(Number(item.lineOrder)) ? Number(item.lineOrder) : index
            }))
          : [],
        paymentEntries: Array.isArray(invoice.paymentEntries)
          ? invoice.paymentEntries.map((payment) => ({
              id: String(payment.id || "").trim(),
              paymentDate: String(payment.paymentDate || "").trim(),
              amount: Number(payment.amount || 0),
              note: String(payment.note || "").trim(),
              createdAt: String(payment.createdAt || "").trim()
            }))
          : []
      }))
    : [];
}

function saveState(nextState) {
  const persistedUi = {
    ...nextState.ui,
    authPanelOpen: false,
    authPanelMode: "profile"
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    view: nextState.view,
    ui: persistedUi
  }));
}

function applyServerState(payload) {
  state.users = payload.users || [];
  state.shows = payload.shows || [];
  state.clients = payload.clients || [];
  state.invoices = payload.invoices || [];
  state.currentUserId = payload.currentUserId || null;
  state.google = payload.google || seedState().google;
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
      shows: state.shows,
      clients: state.clients
    })
  });
  applyServerState(payload);
}

let state = loadLocalUiState();
let mediaThemeListenerAttached = false;
let googleAutoRefreshIntervalId = null;

function getSystemTheme() {
  if (!window.matchMedia) return "dark";
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function applyThemeFromState() {
  const preference = state.ui?.themePreference || "system";
  const resolved = preference === "system" ? getSystemTheme() : preference;
  document.documentElement.dataset.theme = resolved;
}

function setupGoogleAutoRefresh() {
  if (googleAutoRefreshIntervalId) {
    window.clearInterval(googleAutoRefreshIntervalId);
    googleAutoRefreshIntervalId = null;
  }
  const user = getCurrentUser();
  if (!user || user.role !== "admin" || !state.google?.connected) {
    return;
  }
  googleAutoRefreshIntervalId = window.setInterval(async () => {
    if (document.hidden) return;
    try {
      await refreshFromServer();
      render();
    } catch (error) {
      console.warn("Google auto-refresh failed", error);
    }
  }, 1000 * 45);
}

function attachThemeListener() {
  if (mediaThemeListenerAttached || !window.matchMedia) return;
  const media = window.matchMedia("(prefers-color-scheme: dark)");
  const handler = () => {
    if (state.ui?.themePreference === "system") {
      applyThemeFromState();
    }
  };
  if (media.addEventListener) {
    media.addEventListener("change", handler);
  } else {
    media.addListener(handler);
  }
  mediaThemeListenerAttached = true;
}

function getCurrentUser() {
  return state.users.find((user) => user.id === state.currentUserId) || null;
}

function isAdmin(user) {
  return user?.role === "admin";
}

function isAccounts(user) {
  return user?.role === "accounts";
}

function canAccessInvoices(user) {
  return isAdmin(user) || isAccounts(user);
}

function getRoleLabel(role) {
  if (role === "viewer") return "View Only";
  if (role === "accounts") return "Accounts";
  if (role === "admin") return "Admin";
  return "Crew";
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

function getApprovedAccountsUsers() {
  return state.users.filter((user) => user.role === "accounts" && user.approved);
}

function getUserById(id) {
  return state.users.find((user) => user.id === id);
}

function getAssignmentCrewName(assignment) {
  return String(getUserById(assignment?.crewId)?.name || assignment?.manualCrewName || "").trim();
}

function ensureUiState() {
  if (!state.ui) {
    state.ui = {
      editingShowId: null,
      editingInvoiceId: null,
      activeSidebarTab: "calendarPanel",
      activeShowMonth: null,
      selectedShowYear: "all",
      showSortMode: "date",
      selectedCrewFilter: "all",
      googleEntriesView: "needsCompletion",
      invoiceSearchQuery: "",
      invoiceStatusFilter: "all",
      invoicePaymentFilter: "all",
      invoiceClientFilter: "all",
      invoiceSortMode: "issueDate",
      invoiceDraftShowIds: [],
      paymentReconSearchQuery: "",
      paymentReconYear: "all",
      paymentReconMonth: "all",
      paymentReconClient: "all",
      payoutSearchQuery: "",
      payoutYear: "all",
      payoutMonth: "all",
      payoutCrew: "all",
      payoutClient: "all",
      documentShowsYear: "all",
      documentShowsMonth: "all",
      documentInvoicesYear: "all",
      documentInvoicesMonth: "all",
      documentInvoicesClient: "all",
      documentClientFinancialYear: "all",
      documentClientFinancialMonth: "all",
      documentClientFinancialClient: "all",
      documentLedgerClient: "all",
      documentClientsGstFilter: "all",
      documentPayoutYear: "all",
      documentPayoutMonth: "all",
      documentPayoutCrew: "all",
      documentPayoutClient: "all",
      expandedCalendarWeeks: {},
      selectedCalendarShowId: null,
      calendarReturnMode: "month",
      themePreference: "system",
      authPanelMode: "profile",
      authPanelOpen: false
    };
  }

  if (!state.ui.activeSidebarTab) {
    state.ui.activeSidebarTab = "calendarPanel";
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "editingInvoiceId")) {
    state.ui.editingInvoiceId = null;
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "newShowDate")) {
    state.ui.newShowDate = "";
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "showDraftTemplate")) {
    state.ui.showDraftTemplate = null;
  }

  if (!state.ui.showSubtab) {
    state.ui.showSubtab = "list";
  }

  if (!state.ui.showTimelineMode) {
    state.ui.showTimelineMode = "active";
  }

  if (!state.ui.showReturnTab) {
    state.ui.showReturnTab = "showsPanel";
  }

  if (!Number.isFinite(Number(state.ui.showReturnScrollY)) || Number(state.ui.showReturnScrollY) < 0) {
    state.ui.showReturnScrollY = 0;
  }

  if (!state.ui.showSortMode) {
    state.ui.showSortMode = "date";
  }

  if (typeof state.ui.showSearchQuery !== "string") {
    state.ui.showSearchQuery = "";
  }

  if (!Number.isFinite(Number(state.ui.showsPage)) || Number(state.ui.showsPage) < 1) {
    state.ui.showsPage = 1;
  }
  if (!Number.isFinite(Number(state.ui.showsPageSize)) || Number(state.ui.showsPageSize) < 1) {
    state.ui.showsPageSize = 10;
  }

  if (!state.ui.selectedShowYear) {
    state.ui.selectedShowYear = "all";
  }

  if (!state.ui.selectedCrewFilter) {
    state.ui.selectedCrewFilter = "all";
  }

  if (!state.ui.googleEntriesView) {
    state.ui.googleEntriesView = "needsCompletion";
  }

  if (!Number.isFinite(Number(state.ui.googleEntriesPage)) || Number(state.ui.googleEntriesPage) < 1) {
    state.ui.googleEntriesPage = 1;
  }
  if (!Number.isFinite(Number(state.ui.googleEntriesPageSize)) || Number(state.ui.googleEntriesPageSize) < 1) {
    state.ui.googleEntriesPageSize = 10;
  }

  if (!state.ui.invoiceSearchQuery) {
    state.ui.invoiceSearchQuery = "";
  }

  if (!state.ui.invoiceStatusFilter) {
    state.ui.invoiceStatusFilter = "all";
  }

  if (!state.ui.invoicePaymentFilter) {
    state.ui.invoicePaymentFilter = "all";
  }

  if (!state.ui.invoiceClientFilter) {
    state.ui.invoiceClientFilter = "all";
  }

  if (!state.ui.invoiceExportYear) {
    state.ui.invoiceExportYear = "all";
  }

  if (!state.ui.invoiceExportMonth) {
    state.ui.invoiceExportMonth = "all";
  }

  if (!state.ui.invoiceLightDesignerFilter) {
    state.ui.invoiceLightDesignerFilter = "all";
  }

  if (!Array.isArray(state.ui.invoiceExportColumns) || !state.ui.invoiceExportColumns.length) {
    state.ui.invoiceExportColumns = DEFAULT_INVOICE_EXPORT_COLUMN_KEYS;
  }
  state.ui.invoiceExportColumns = state.ui.invoiceExportColumns.filter((key) => INVOICE_EXPORT_COLUMNS.some((column) => column.key === key));
  if (!state.ui.invoiceExportColumns.length) {
    state.ui.invoiceExportColumns = DEFAULT_INVOICE_EXPORT_COLUMN_KEYS;
  }

  if (!state.ui.invoiceSortMode) {
    state.ui.invoiceSortMode = "issueDate";
  }

  if (!Number.isFinite(Number(state.ui.invoiceRegisterPage)) || Number(state.ui.invoiceRegisterPage) < 1) {
    state.ui.invoiceRegisterPage = 1;
  }
  if (!Number.isFinite(Number(state.ui.invoiceRegisterPageSize)) || Number(state.ui.invoiceRegisterPageSize) < 1) {
    state.ui.invoiceRegisterPageSize = 10;
  }

  if (!Array.isArray(state.ui.invoiceDraftShowIds)) {
    state.ui.invoiceDraftShowIds = [];
  }

  if (typeof state.ui.paymentReconSearchQuery !== "string") {
    state.ui.paymentReconSearchQuery = "";
  }

  if (!state.ui.paymentReconYear) {
    state.ui.paymentReconYear = "all";
  }

  if (!state.ui.paymentReconMonth) {
    state.ui.paymentReconMonth = "all";
  }

  if (!state.ui.paymentReconClient) {
    state.ui.paymentReconClient = "all";
  }

  if (!Number.isFinite(Number(state.ui.paymentReconPage)) || Number(state.ui.paymentReconPage) < 1) {
    state.ui.paymentReconPage = 1;
  }

  if (!Number.isFinite(Number(state.ui.paymentReconPageSize)) || Number(state.ui.paymentReconPageSize) < 1) {
    state.ui.paymentReconPageSize = 10;
  }

  if (typeof state.ui.payoutSearchQuery !== "string") {
    state.ui.payoutSearchQuery = "";
  }

  if (!state.ui.payoutYear) {
    state.ui.payoutYear = "all";
  }

  if (!state.ui.payoutMonth) {
    state.ui.payoutMonth = "all";
  }

  if (!state.ui.payoutCrew) {
    state.ui.payoutCrew = "all";
  }

  if (!state.ui.payoutClient) {
    state.ui.payoutClient = "all";
  }

  if (!Number.isFinite(Number(state.ui.payoutPage)) || Number(state.ui.payoutPage) < 1) {
    state.ui.payoutPage = 1;
  }

  if (!Number.isFinite(Number(state.ui.payoutPageSize)) || Number(state.ui.payoutPageSize) < 1) {
    state.ui.payoutPageSize = 10;
  }

  if (!state.ui.documentShowsYear) state.ui.documentShowsYear = "all";
  if (!state.ui.documentShowsMonth) state.ui.documentShowsMonth = "all";
  if (!state.ui.documentInvoicesYear) state.ui.documentInvoicesYear = "all";
  if (!state.ui.documentInvoicesMonth) state.ui.documentInvoicesMonth = "all";
  if (!state.ui.documentInvoicesClient) state.ui.documentInvoicesClient = "all";
  if (!state.ui.documentClientFinancialYear) state.ui.documentClientFinancialYear = "all";
  if (!state.ui.documentClientFinancialMonth) state.ui.documentClientFinancialMonth = "all";
  if (!state.ui.documentClientFinancialClient) state.ui.documentClientFinancialClient = "all";
  if (!state.ui.documentLedgerClient) state.ui.documentLedgerClient = "all";
  if (!state.ui.documentClientsGstFilter) state.ui.documentClientsGstFilter = "all";
  if (!state.ui.documentPayoutYear) state.ui.documentPayoutYear = "all";
  if (!state.ui.documentPayoutMonth) state.ui.documentPayoutMonth = "all";
  if (!state.ui.documentPayoutCrew) state.ui.documentPayoutCrew = "all";
  if (!state.ui.documentPayoutClient) state.ui.documentPayoutClient = "all";

  if (!state.ui.invoiceDraftTemplate || typeof state.ui.invoiceDraftTemplate !== "object") {
    state.ui.invoiceDraftTemplate = null;
  }

  if (!state.ui.invoiceSubtab) {
    state.ui.invoiceSubtab = "create";
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "markPaymentInvoiceId")) {
    state.ui.markPaymentInvoiceId = null;
  }

  if (!Number.isFinite(Number(state.ui.clientsPage)) || Number(state.ui.clientsPage) < 1) {
    state.ui.clientsPage = 1;
  }
  if (!Number.isFinite(Number(state.ui.clientsPageSize)) || Number(state.ui.clientsPageSize) < 1) {
    state.ui.clientsPageSize = 10;
  }

  if (!state.ui.clientsSubtab) {
    state.ui.clientsSubtab = "list";
  }
  if (!Object.prototype.hasOwnProperty.call(state.ui, "selectedClientDetailId")) {
    state.ui.selectedClientDetailId = null;
  }
  if (typeof state.ui.clientSearchQuery !== "string") {
    state.ui.clientSearchQuery = "";
  }
  if (!state.ui.clientGstFilter) {
    state.ui.clientGstFilter = "all";
  }
  if (!state.ui.clientExportYear) {
    state.ui.clientExportYear = "all";
  }
  if (!state.ui.clientExportMonth) {
    state.ui.clientExportMonth = "all";
  }
  if (!state.ui.clientExportClientId) {
    state.ui.clientExportClientId = "all";
  }

  if (!Object.prototype.hasOwnProperty.call(state.ui, "dirtyForm")) {
    state.ui.dirtyForm = null;
  }

  const normalizedInvoicePrintCopies = Number(state.ui.invoicePrintCopies || 1);
  state.ui.invoicePrintCopies = [1, 2, 3, 4, 5].includes(normalizedInvoicePrintCopies) ? normalizedInvoicePrintCopies : 1;

  if (!state.ui.invoiceLineDefaults || typeof state.ui.invoiceLineDefaults !== "object") {
    state.ui.invoiceLineDefaults = { description: "", sac: "" };
  }
  state.ui.invoiceLineDefaults.description = String(state.ui.invoiceLineDefaults.description || "");
  state.ui.invoiceLineDefaults.sac = String(state.ui.invoiceLineDefaults.sac || "");

  if (!state.ui.googleArchiveYear) {
    state.ui.googleArchiveYear = "all";
  }

  if (!state.ui.googleArchiveMonth) {
    state.ui.googleArchiveMonth = "all";
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

  if (!state.ui.themePreference) {
    state.ui.themePreference = "system";
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
  state.ui.newShowDate = "";
  state.ui.showDraftTemplate = null;
}

function getShowReturnTab() {
  ensureUiState();
  const returnTab = state.ui.showReturnTab || "showsPanel";
  return ["calendarPanel", "googleEntriesPanel", "showsPanel"].includes(returnTab) ? returnTab : "showsPanel";
}

function captureShowReturnContext(returnTab = state.ui.activeSidebarTab || "showsPanel") {
  ensureUiState();
  state.ui.showReturnTab = ["calendarPanel", "googleEntriesPanel", "showsPanel"].includes(returnTab) ? returnTab : "showsPanel";
  state.ui.showReturnScrollY = Math.max(
    window.scrollY || 0,
    document.documentElement?.scrollTop || 0,
    document.body?.scrollTop || 0
  );
}

function restoreShowReturnContext() {
  ensureUiState();
  const scrollY = Number(state.ui.showReturnScrollY || 0);
  window.requestAnimationFrame(() => {
    window.scrollTo({ top: scrollY, behavior: "auto" });
  });
}

function getDirtyFormLabel(formKey = state.ui?.dirtyForm) {
  if (formKey === "show") return "show form";
  if (formKey === "client") return "client form";
  if (formKey === "invoice") return "invoice form";
  return "form";
}

function setDirtyForm(formKey) {
  ensureUiState();
  if (state.ui.dirtyForm === formKey) return;
  state.ui.dirtyForm = formKey;
  saveState(state);
}

function clearDirtyForm(formKey = null) {
  ensureUiState();
  if (formKey && state.ui.dirtyForm !== formKey) return;
  if (!state.ui.dirtyForm) return;
  state.ui.dirtyForm = null;
  saveState(state);
}

function hasDirtyForm(formKey = null) {
  ensureUiState();
  if (!state.ui.dirtyForm) return false;
  return formKey ? state.ui.dirtyForm === formKey : true;
}

function confirmDiscardDirtyForm(targetLabel = "continue") {
  if (!hasDirtyForm()) return true;
  return window.confirm(`You have unsaved changes in the ${getDirtyFormLabel()}. Do you want to discard them and ${targetLabel}?`);
}

function wireDirtyFormTracking(form, formKey) {
  if (!form) return;
  if (form.dataset.dirtyTrackingWired === "true") return;
  form.dataset.dirtyTrackingWired = "true";
  const markDirty = (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.closest("input, textarea, select")) return;
    setDirtyForm(formKey);
  };
  form.addEventListener("input", markDirty);
  form.addEventListener("change", markDirty);
}

function getDirtyTabLabel(baseLabel, formKey) {
  if (!hasDirtyForm(formKey)) return baseLabel;
  return `${baseLabel} <span class="dirty-tab-badge" aria-label="Unsaved changes">Pending</span>`;
}

function resetInvoiceEditingState() {
  ensureUiState();
  state.ui.editingInvoiceId = null;
  state.ui.invoiceDraftShowIds = [];
  state.ui.invoiceDraftTemplate = null;
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
  if (user && isAdmin(user)) {
    return [
      { id: "calendarPanel", label: "Calendar", meta: "Month, week, day" },
      { id: "showsPanel", label: "Shows", meta: "Create and view shows" },
      { id: "invoicesPanel", label: "Invoices", meta: "Create and track billing" },
      { id: "clientsPanel", label: "Clients", meta: "Client master and billing info" },
      { id: "documentsPanel", label: "Document Center", meta: "Exports and ledgers" },
      { id: "googleEntriesPanel", label: "Google Calendar", meta: "Imported and synced shows" },
      { id: "crewAdminPanel", label: "Crew Management", meta: "Add or remove crew" }
    ];
  }

  if (user && isAccounts(user)) {
    return [
      { id: "calendarPanel", label: "Calendar", meta: "Month, week, day" },
      { id: "showsPanel", label: "Shows", meta: "Create and view shows" },
      { id: "invoicesPanel", label: "Invoices", meta: "Billing and collections" },
      { id: "clientsPanel", label: "Clients", meta: "Client master and billing info" },
      { id: "documentsPanel", label: "Document Center", meta: "Exports and ledgers" }
    ];
  }

  return [
    { id: "calendarPanel", label: "Calendar", meta: "Shared schedule" },
    { id: "showsPanel", label: "Shows", meta: "Visible entries" },
    ...(user?.role === "viewer" ? [{ id: "legendPanel", label: "Crew", meta: "Colors and teams" }] : [])
  ];
}

function ensureActiveSidebarTab(user) {
  if (state.ui.activeSidebarTab === "showFormPanel") {
    state.ui.activeSidebarTab = "showsPanel";
    state.ui.showSubtab = "create";
  }
  const tabs = getSidebarTabs(user);
  if (!tabs.some((tab) => tab.id === state.ui.activeSidebarTab)) {
    state.ui.activeSidebarTab = tabs[0].id;
  }
}

function getTakenColors(excludeUserId = null) {
  return state.users
    .filter((user) => (user.role === "crew" || user.role === "admin") && user.approved && user.id !== excludeUserId)
    .map((user) => resolveCrewColor(user.color))
    .filter(Boolean);
}

function visibleShowsForUser(user) {
  if (!user) return [];
  if (user.role === "admin" || user.role === "viewer" || user.role === "accounts") {
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

function getGoogleLinkedShows(shows = state.shows) {
  return shows.filter((show) => show.googleEventId || show.googleSyncSource === "google" || show.googleSyncStatus);
}

function getInvoiceStatusLabel(invoice) {
  const today = dateKey(new Date());
  if (invoice.status === "cancelled") return "Cancelled";
  if (invoice.status === "paid" || Number(invoice.balanceDue || 0) <= 0) return "Paid";
  if (Number(invoice.amountPaid || 0) > 0 && Number(invoice.balanceDue || 0) > 0) return "Partially Paid";
  if (invoice.status === "partially_paid") return "Partially Paid";
  if (invoice.dueDate && invoice.dueDate < today) return "Overdue";
  if (invoice.status === "sent") return "Sent";
  return "Draft";
}

function getInvoiceStatusTone(invoice) {
  const label = getInvoiceStatusLabel(invoice);
  if (label === "Paid") return "paid";
  if (label === "Overdue") return "overdue";
  if (label === "Cancelled") return "cancelled";
  if (label === "Partially Paid") return "partial";
  if (label === "Sent") return "sent";
  return "draft";
}

function getInvoiceLightDesignerLabel(invoice) {
  const names = [...new Set((invoice.lineItems || [])
    .flatMap((item) => parseInvoiceShowIds(item.showId)
      .map((showId) => state.shows.find((show) => show.id === showId))
      .filter(Boolean))
    .flatMap((show) => (show.assignments || [])
      .map((assignment) => getUserById(assignment.lightDesignerId)?.name)
      .filter(Boolean)))].sort((a, b) => a.localeCompare(b));
  return names.join(", ");
}

function getInvoiceLinkedShowNames(invoice) {
  return [...new Set((invoice.lineItems || [])
    .flatMap((item) => parseInvoiceShowIds(item.showId)
      .map((showId) => state.shows.find((show) => show.id === showId)?.showName))
    .filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function getInvoiceClientLabel(invoice) {
  const client = getClientById(invoice?.clientId) || getClientByName(invoice?.clientName);
  return client ? getClientDisplayName(client) : String(invoice?.clientName || "Client TBD").trim() || "Client TBD";
}

function buildInvoiceDraftTemplateFromInvoice(invoice) {
  if (!invoice) return null;
  return {
    invoiceNumber: makeDefaultInvoiceNumber(),
    clientId: invoice.clientId || "",
    clientName: invoice.clientName || "",
    issueDate: dateKey(new Date()),
    dueDate: "",
    status: "draft",
    amountPaid: 0,
    notes: invoice.notes || "",
    details: {
      ...normalizeInvoiceDetails(invoice.details),
      paymentTerms: normalizeInvoiceDetails(invoice.details).paymentTerms || "Net 15"
    },
    lineItems: (invoice.lineItems || []).map((item) => ({
      id: uid("line"),
      showId: item.showId || "",
      description: item.description || "",
      sac: item.sac || item.hsnSac || "",
      customDetails: item.customDetails || item.custom_details || "",
      discount: item.discount || "",
      quantity: Number(item.quantity || 1),
      unitRate: Number(item.unitRate || 0),
      amount: Number(item.amount || 0)
    }))
  };
}

function getInvoicePaymentBucket(invoice) {
  const label = getInvoiceStatusLabel(invoice);
  if (label === "Paid") return "paid";
  if (label === "Partially Paid") return "partiallyPaid";
  return "unpaid";
}

function normalizeInvoiceDetails(details = {}) {
  const source = details && typeof details === "object" ? details : {};
  return {
    companyName: String(source.companyName || "PixelBug").trim(),
    companyAddress: String(source.companyAddress || "").trim(),
    companyEmail: String(source.companyEmail || "").trim(),
    companyPhone: String(source.companyPhone || "").trim(),
    companyGstin: String(source.companyGstin || "").trim(),
    clientBillingAddress: String(source.clientBillingAddress || "").trim(),
    clientGstin: String(source.clientGstin || "").trim(),
    placeOfSupply: String(source.placeOfSupply || "").trim(),
    paymentTerms: String(source.paymentTerms || "Net 15").trim(),
    bankAccountName: String(source.bankAccountName || "").trim(),
    bankName: String(source.bankName || "").trim(),
    bankAccountNumber: String(source.bankAccountNumber || "").trim(),
    bankIfsc: String(source.bankIfsc || "").trim(),
    footerNote: String(source.footerNote || "Please include the invoice number with your payment reference.").trim()
  };
}

function getDueDateFromTerms(issueDate, paymentTerms) {
  const normalizedIssueDate = String(issueDate || "").trim();
  if (!normalizedIssueDate) return "";
  const normalizedTerms = String(paymentTerms || "Net 15").trim().toLowerCase();
  const offsets = {
    "due on receipt": 0,
    "net 10": 10,
    "net 15": 15,
    "net 30": 30
  };
  const offsetDays = offsets[normalizedTerms] ?? 0;
  return dateKey(addDays(parseDateKey(normalizedIssueDate), offsetDays));
}

function getDiscountAmount(rawDiscount, baseAmount) {
  const normalized = String(rawDiscount || "").trim();
  if (!normalized) return 0;
  const numeric = Number(normalized.replace(/%$/, "").trim());
  if (!Number.isFinite(numeric) || numeric <= 0) return 0;
  const discountAmount = normalized.endsWith("%")
    ? Number(baseAmount || 0) * (numeric / 100)
    : numeric;
  return Math.round(Math.min(Number(baseAmount || 0), Math.max(0, discountAmount)) * 100) / 100;
}

function isMaharashtraSupply(placeOfSupply) {
  return String(placeOfSupply || "").toLowerCase().includes("maharashtra");
}

function getInvoiceCalculationFromValues(lineItems = [], placeOfSupply = "") {
  const grossSubtotal = lineItems.reduce((sum, item) => sum + (Number(item.quantity || 1) * Number(item.unitRate || 0)), 0);
  const discountAmount = lineItems.reduce((sum, item) => {
    const baseAmount = Number(item.quantity || 1) * Number(item.unitRate || 0);
    return sum + getDiscountAmount(item.discount || "", baseAmount);
  }, 0);
  const taxableAmount = Math.max(0, grossSubtotal - discountAmount);
  const intraState = isMaharashtraSupply(placeOfSupply);
  const sgstAmount = intraState ? taxableAmount * 0.09 : 0;
  const cgstAmount = intraState ? taxableAmount * 0.09 : 0;
  const igstAmount = intraState ? 0 : taxableAmount * 0.18;
  const taxAmount = sgstAmount + cgstAmount + igstAmount;
  return {
    grossSubtotal: Math.round(grossSubtotal * 100) / 100,
    discountAmount: Math.round(discountAmount * 100) / 100,
    taxableAmount: Math.round(taxableAmount * 100) / 100,
    sgstAmount: Math.round(sgstAmount * 100) / 100,
    cgstAmount: Math.round(cgstAmount * 100) / 100,
    igstAmount: Math.round(igstAmount * 100) / 100,
    taxAmount: Math.round(taxAmount * 100) / 100,
    totalAmount: Math.round((taxableAmount + taxAmount) * 100) / 100
  };
}

function getFixedInvoiceCompanyProfile() {
  return {
    name: "PixelBug",
    gstin: "27AAZFP6374P1ZW",
    email: "pixelbugsolutions@gmail.com",
    phone: "+91 7666426289",
    address: "Pune, Maharashtra, India - 411009",
    website: "www.pixelbug.in",
    bankAccountName: "PixelBug",
    bankName: "HDFC Bank",
    bankAccountNumber: "50200055939716",
    bankAccountType: "Current Account",
    bankBranchName: "Karvenagar, Pune",
    bankIfsc: "HDFC0001115",
    signatureImage: "signature-sachin.jpg",
    signatureHolder: "Sachin Dunakhe, Director, PixelBug",
    footerNote: "Please include the invoice number with your payment reference."
  };
}

const INDIAN_STATE_OPTIONS = [
  "Andaman and Nicobar Islands (35)",
  "Andhra Pradesh (37)",
  "Arunachal Pradesh (12)",
  "Assam (18)",
  "Bihar (10)",
  "Chandigarh (04)",
  "Chhattisgarh (22)",
  "Dadra and Nagar Haveli and Daman and Diu (26)",
  "Delhi (07)",
  "Goa (30)",
  "Gujarat (24)",
  "Haryana (06)",
  "Himachal Pradesh (02)",
  "Jammu and Kashmir (01)",
  "Jharkhand (20)",
  "Karnataka (29)",
  "Kerala (32)",
  "Ladakh (38)",
  "Lakshadweep (31)",
  "Madhya Pradesh (23)",
  "Maharashtra (27)",
  "Manipur (14)",
  "Meghalaya (17)",
  "Mizoram (15)",
  "Nagaland (13)",
  "Odisha (21)",
  "Puducherry (34)",
  "Punjab (03)",
  "Rajasthan (08)",
  "Sikkim (11)",
  "Tamil Nadu (33)",
  "Telangana (36)",
  "Tripura (16)",
  "Uttar Pradesh (09)",
  "Uttarakhand (05)",
  "West Bengal (19)"
];

const PAGINATION_PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

function getPaginationSlice(items = [], pageKey, pageSizeKey) {
  const totalItems = items.length;
  let pageSize = Number(state.ui[pageSizeKey] || 10);
  if (!PAGINATION_PAGE_SIZE_OPTIONS.includes(pageSize)) {
    pageSize = 10;
  }
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  let currentPage = Math.max(1, Number(state.ui[pageKey] || 1));
  currentPage = Math.min(currentPage, totalPages);
  state.ui[pageKey] = currentPage;
  state.ui[pageSizeKey] = pageSize;
  const startIndex = (currentPage - 1) * pageSize;
  return {
    items: items.slice(startIndex, startIndex + pageSize),
    totalItems,
    pageSize,
    currentPage,
    totalPages,
    startItem: totalItems ? startIndex + 1 : 0,
    endItem: Math.min(totalItems, startIndex + pageSize)
  };
}

function renderPaginationControls(id, pagination, label = "entries") {
  if (!pagination.totalItems) return "";
  const visiblePages = getPaginationRange(pagination.currentPage, pagination.totalPages);
  return `
    <div class="pagination-bar" data-pagination="${id}">
      <div class="pagination-info">
        <strong>${pagination.startItem}-${pagination.endItem}</strong>
        <span>of ${pagination.totalItems} ${label}</span>
      </div>
      <div class="pagination-controls">
        <label class="sort-control pagination-size-control">
          <span>Per Page</span>
          <select data-pagination-size="${id}">
            ${PAGINATION_PAGE_SIZE_OPTIONS.map((size) => `<option value="${size}" ${pagination.pageSize === size ? "selected" : ""}>${size}</option>`).join("")}
          </select>
        </label>
        <div class="pagination-nav">
          <button type="button" class="ghost small" data-pagination-page="${id}" data-page-target="first" ${pagination.currentPage <= 1 ? "disabled" : ""}>First</button>
          <button type="button" class="ghost small" data-pagination-page="${id}" data-page-target="prev" ${pagination.currentPage <= 1 ? "disabled" : ""}>Prev</button>
          <div class="pagination-page-range" aria-label="Page range">
            ${visiblePages.map((page) => page === "ellipsis"
              ? '<span class="pagination-ellipsis">...</span>'
              : `<button type="button" class="ghost small ${pagination.currentPage === page ? "is-active" : ""}" data-pagination-page="${id}" data-page-target="${page}" aria-current="${pagination.currentPage === page ? "page" : "false"}">${page}</button>`
            ).join("")}
          </div>
          <button type="button" class="ghost small" data-pagination-page="${id}" data-page-target="next" ${pagination.currentPage >= pagination.totalPages ? "disabled" : ""}>Next</button>
          <button type="button" class="ghost small" data-pagination-page="${id}" data-page-target="last" ${pagination.currentPage >= pagination.totalPages ? "disabled" : ""}>Last</button>
        </div>
      </div>
    </div>
  `;
}

function getPaginationRange(currentPage, totalPages) {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }
  const pages = new Set([1, totalPages, currentPage - 1, currentPage, currentPage + 1].filter((page) => page >= 1 && page <= totalPages));
  const sortedPages = [...pages].sort((a, b) => a - b);
  return sortedPages.flatMap((page, index) => {
    const previous = sortedPages[index - 1];
    if (index > 0 && page - previous > 1) {
      return ["ellipsis", page];
    }
    return [page];
  });
}

function getStateFromGstin(gstin) {
  const code = String(gstin || "").trim().slice(0, 2);
  if (!/^\d{2}$/.test(code)) return "";
  return INDIAN_STATE_OPTIONS.find((stateName) => stateName.endsWith(`(${code})`)) || "";
}

function wirePaginationControls(root, id, pageKey, pageSizeKey, renderCallback) {
  root.querySelector(`[data-pagination-size="${id}"]`)?.addEventListener("change", (event) => {
    state.ui[pageSizeKey] = Number(event.currentTarget.value || 10);
    state.ui[pageKey] = 1;
    saveState(state);
    renderCallback();
  });
  root.querySelectorAll(`[data-pagination-page="${id}"]`).forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.pageTarget;
      const currentPage = Math.max(1, Number(state.ui[pageKey] || 1));
      if (target === "first") {
        state.ui[pageKey] = 1;
      } else if (target === "last") {
        state.ui[pageKey] = Number.MAX_SAFE_INTEGER;
      } else if (target === "prev") {
        state.ui[pageKey] = Math.max(1, currentPage - 1);
      } else if (target === "next") {
        state.ui[pageKey] = currentPage + 1;
      } else {
        state.ui[pageKey] = Math.max(1, Number(target || 1));
      }
      saveState(state);
      renderCallback();
    });
  });
}

function normalizeClient(client = {}) {
  return {
    id: String(client.id || uid("client")).trim(),
    name: String(client.name || "").trim(),
    state: String(client.state || "").trim(),
    billingAddress: String(client.billingAddress || "").trim(),
    gstin: String(client.gstin || "").trim(),
    contactName: String(client.contactName || "").trim(),
    contactEmail: String(client.contactEmail || "").trim().toLowerCase(),
    contactPhone: String(client.contactPhone || "").trim(),
    notes: String(client.notes || "").trim(),
    createdAt: String(client.createdAt || new Date().toISOString()).trim(),
    updatedAt: String(client.updatedAt || new Date().toISOString()).trim()
  };
}

function getClientDisplayName(client, clients = state.clients || []) {
  const normalizedClient = normalizeClient(client);
  const normalizedName = normalizedClient.name.trim().toLowerCase();
  if (!normalizedName) return "";
  const duplicateCount = (clients || [])
    .map((item) => normalizeClient(item))
    .filter((item) => item.name.trim().toLowerCase() === normalizedName)
    .length;
  return duplicateCount > 1 && normalizedClient.state
    ? `${normalizedClient.name} (${normalizedClient.state})`
    : normalizedClient.name;
}

function getSortedClients() {
  return [...(state.clients || [])]
    .map((client) => normalizeClient(client))
    .filter((client) => client.name)
    .sort((a, b) => getClientDisplayName(a).localeCompare(getClientDisplayName(b)));
}

function getClientById(clientId) {
  if (!clientId) return null;
  return (state.clients || []).find((client) => client.id === clientId) || null;
}

function getClientByName(name) {
  const normalized = String(name || "").trim().toLowerCase();
  if (!normalized) return null;
  return (state.clients || []).find((client) => {
    const normalizedClient = normalizeClient(client);
    return normalizedClient.name.trim().toLowerCase() === normalized
      || getClientDisplayName(normalizedClient).trim().toLowerCase() === normalized;
  }) || null;
}

function getClientDisplayValue(clientId, fallbackName = "") {
  const client = getClientById(clientId) || getClientByName(fallbackName);
  return client ? getClientDisplayName(client) : String(fallbackName || "").trim();
}

function sortInvoices(invoices = state.invoices) {
  const mode = state.ui.invoiceSortMode || "issueDate";
  const items = [...invoices];
  if (mode === "dueDate") {
    return items.sort((a, b) =>
      String(a.dueDate || "9999-12-31").localeCompare(String(b.dueDate || "9999-12-31"))
      || String(b.issueDate || "").localeCompare(String(a.issueDate || ""))
      || String(a.invoiceNumber || "").localeCompare(String(b.invoiceNumber || ""))
    );
  }
  if (mode === "client") {
    return items.sort((a, b) =>
      String(a.clientName || "").localeCompare(String(b.clientName || ""))
      || String(a.invoiceNumber || "").localeCompare(String(b.invoiceNumber || ""))
    );
  }
  if (mode === "lightDesigner") {
    return items.sort((a, b) =>
      getInvoiceLightDesignerLabel(a).localeCompare(getInvoiceLightDesignerLabel(b))
      || String(a.clientName || "").localeCompare(String(b.clientName || ""))
      || String(a.invoiceNumber || "").localeCompare(String(b.invoiceNumber || ""))
    );
  }
  return items.sort((a, b) => {
    const issueCompare = String(b.issueDate || "").localeCompare(String(a.issueDate || ""));
    if (issueCompare !== 0) return issueCompare;
    return String(b.invoiceNumber || "").localeCompare(String(a.invoiceNumber || ""));
  });
}

function getInvoiceClientOptions(invoices = state.invoices) {
  return [...new Set(invoices.map((invoice) => String(invoice.clientName || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function getInvoiceYearOptions(invoices = state.invoices) {
  return [...new Set(invoices.map((invoice) => String(invoice.issueDate || "").slice(0, 4)).filter(Boolean))].sort((a, b) => b.localeCompare(a));
}

function getInvoiceMonthOptions(invoices = state.invoices, selectedYear = "all") {
  const monthKeys = invoices
    .filter((invoice) => selectedYear === "all" || String(invoice.issueDate || "").slice(0, 4) === selectedYear)
    .map((invoice) => String(invoice.issueDate || "").slice(0, 7))
    .filter(Boolean);
  return [...new Set(monthKeys)].sort((a, b) => b.localeCompare(a));
}

function getInvoiceLightDesignerOptions(invoices = state.invoices) {
  return [...new Set(invoices.flatMap((invoice) => getInvoiceLightDesignerLabel(invoice).split(",").map((name) => name.trim()).filter(Boolean)))]
    .sort((a, b) => a.localeCompare(b));
}

function getCrewPayoutRows(shows = state.shows) {
  return [...shows]
    .flatMap((show) => (show.assignments || []).map((assignment, index) => {
      const crewName = getAssignmentCrewName(assignment);
      const operatorAmount = Math.round(Number(assignment.operatorAmount || 0) * 100) / 100;
      if (!crewName || operatorAmount <= 0) return null;
      return {
        id: `${show.id}:${assignment.crewId || "manual"}:${index}`,
        showId: show.id,
        showName: show.showName || "Untitled Show",
        showDate: getShowStartDate(show),
        showDateLabel: formatDateRange(getShowStartDate(show), getShowEndDate(show)),
        clientLabel: getClientDisplayValue(show.clientId, show.client),
        clientId: show.clientId || getClientByName(show.client)?.id || "",
        crewId: assignment.crewId || "",
        crewName,
        lightDesigner: getUserById(assignment.lightDesignerId)?.name || "",
        location: show.location || show.venue || "",
        operatorAmount,
        status: show.showStatus || "confirmed",
        statusLabel: show.showStatus === "tentative" ? "Tentative" : "Confirmed",
        notes: String(assignment.notes || "").trim()
      };
    }))
    .filter(Boolean)
    .sort((a, b) =>
      String(b.showDate || "").localeCompare(String(a.showDate || ""))
      || String(a.showName || "").localeCompare(String(b.showName || ""))
      || String(a.crewName || "").localeCompare(String(b.crewName || ""))
    );
}

function getCrewPayoutYearOptions(rows = getCrewPayoutRows()) {
  return [...new Set(rows.map((row) => String(row.showDate || "").slice(0, 4)).filter(Boolean))]
    .sort((a, b) => b.localeCompare(a));
}

function getCrewPayoutMonthOptions(rows = getCrewPayoutRows(), selectedYear = "all") {
  return [...new Set(rows
    .filter((row) => selectedYear === "all" || String(row.showDate || "").slice(0, 4) === selectedYear)
    .map((row) => String(row.showDate || "").slice(0, 7))
    .filter(Boolean))]
    .sort((a, b) => b.localeCompare(a));
}

function filterCrewPayoutRows(rows = getCrewPayoutRows()) {
  const query = String(state.ui.payoutSearchQuery || "").trim().toLowerCase();
  const year = state.ui.payoutYear || "all";
  const month = state.ui.payoutMonth || "all";
  const crew = state.ui.payoutCrew || "all";
  const client = state.ui.payoutClient || "all";
  return rows.filter((row) => {
    const rowYear = String(row.showDate || "").slice(0, 4);
    const rowMonth = String(row.showDate || "").slice(0, 7);
    const haystack = [
      row.showName,
      row.clientLabel,
      row.crewName,
      row.lightDesigner,
      row.location,
      row.notes,
      row.statusLabel
    ].join(" ").toLowerCase();
    const matchesQuery = !query || haystack.includes(query);
    const matchesYear = year === "all" || rowYear === year;
    const matchesMonth = month === "all" || rowMonth === month;
    const matchesCrew = crew === "all" || row.crewName === crew;
    const matchesClient = client === "all" || row.clientLabel === client;
    return matchesQuery && matchesYear && matchesMonth && matchesCrew && matchesClient;
  });
}

function getFilteredClientsList() {
  const clients = getSortedClients();
  const clientSearchQuery = String(state.ui.clientSearchQuery || "").trim().toLowerCase();
  const clientGstFilter = state.ui.clientGstFilter || "all";
  return clients.filter((client) => {
    const matchesSearch = !clientSearchQuery || [
      getClientDisplayName(client),
      client.name,
      client.state,
      client.gstin,
      client.contactName,
      client.contactEmail,
      client.contactPhone,
      client.billingAddress,
      client.notes
    ].some((value) => String(value || "").toLowerCase().includes(clientSearchQuery));
    if (!matchesSearch) return false;
    if (clientGstFilter === "missing") {
      return !String(client.gstin || "").trim();
    }
    if (clientGstFilter === "available") {
      return Boolean(String(client.gstin || "").trim());
    }
    return true;
  });
}

function getFilteredShowsForDocuments(user = getCurrentUser()) {
  const shows = filterShowsBySelectedCrew(visibleShowsForUser(user));
  const showSearchQuery = String(state.ui.showSearchQuery || "").trim().toLowerCase();
  const todayKey = dateKey(new Date());
  const filteredShows = shows.filter((show) => {
    const showMonthKey = getShowStartDate(show).slice(0, 7);
    const yearMatch = state.ui.selectedShowYear === "all" || showMonthKey.startsWith(state.ui.selectedShowYear);
    const monthMatch = state.ui.activeShowMonth === "all" || showMonthKey === state.ui.activeShowMonth;
    const searchableValues = [
      show.showName,
      show.client,
      show.location,
      show.venue,
      ...show.assignments.map((assignment) => getAssignmentCrewName(assignment))
    ];
    const searchMatch = !showSearchQuery || searchableValues.some((value) => String(value || "").toLowerCase().includes(showSearchQuery));
    return yearMatch && monthMatch && searchMatch;
  });
  const currentShows = filteredShows.filter((show) => getShowStartDate(show) <= todayKey && getShowEndDate(show) >= todayKey);
  const upcomingShows = filteredShows.filter((show) => getShowStartDate(show) > todayKey);
  const pastShows = filteredShows.filter((show) => getShowEndDate(show) < todayKey);
  const timelineMode = state.ui.showTimelineMode || "active";
  return sortShows(timelineMode === "past" ? pastShows : [...currentShows, ...upcomingShows], state.ui.showSortMode);
}

function filterInvoices(invoices = state.invoices) {
  const query = String(state.ui.invoiceSearchQuery || "").trim().toLowerCase();
  const statusFilter = state.ui.invoiceStatusFilter || "all";
  const paymentFilter = state.ui.invoicePaymentFilter || "all";
  const clientFilter = state.ui.invoiceClientFilter || "all";
  const exportYear = state.ui.invoiceExportYear || "all";
  const exportMonth = state.ui.invoiceExportMonth || "all";
  const lightDesignerFilter = state.ui.invoiceLightDesignerFilter || "all";
  return invoices.filter((invoice) => {
    const statusLabel = getInvoiceStatusLabel(invoice).toLowerCase();
    const paymentBucket = getInvoicePaymentBucket(invoice);
    const clientName = String(invoice.clientName || "").trim();
    const invoiceYear = String(invoice.issueDate || "").slice(0, 4);
    const invoiceMonth = String(invoice.issueDate || "").slice(0, 7);
    const lightDesignerNames = getInvoiceLightDesignerLabel(invoice).split(",").map((name) => name.trim()).filter(Boolean);
    const haystack = [
      invoice.invoiceNumber,
      invoice.clientName,
      statusLabel,
      paymentBucket,
      getInvoiceLightDesignerLabel(invoice),
      ...(invoice.lineItems || []).map((item) => item.description),
      ...(invoice.lineItems || []).map((item) => item.sac),
      ...(invoice.lineItems || []).map((item) => item.customDetails || item.custom_details),
      ...(invoice.lineItems || []).map((item) => getInvoiceLineShowLabel(item.showId))
    ].join(" ").toLowerCase();
    const matchesQuery = !query || haystack.includes(query);
    const matchesStatus = statusFilter === "all" || statusLabel === statusFilter;
    const matchesPayment = paymentFilter === "all" || paymentBucket === paymentFilter;
    const matchesClient = clientFilter === "all" || clientName === clientFilter;
    const matchesYear = exportYear === "all" || invoiceYear === exportYear;
    const matchesMonth = exportMonth === "all" || invoiceMonth === exportMonth;
    const matchesLightDesigner = lightDesignerFilter === "all" || lightDesignerNames.includes(lightDesignerFilter);
    return matchesQuery && matchesStatus && matchesPayment && matchesClient && matchesYear && matchesMonth && matchesLightDesigner;
  });
}

function getShowsLinkedToInvoices() {
  return new Set((state.invoices || []).flatMap((invoice) => (invoice.lineItems || []).flatMap((item) => parseInvoiceShowIds(item.showId))));
}

function getInvoicePaymentHistoryMarkup(invoice) {
  const payments = Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : [];
  if (!payments.length) return "";
  return `
    <div class="invoice-payment-history">
      <strong>Payment History</strong>
      ${payments.map((payment) => `
        <div class="invoice-payment-history-row">
          <span>${escapeHtml(formatInvoiceDate(payment.paymentDate))}</span>
          <strong>${formatCurrency(payment.amount)}</strong>
          ${payment.note ? `<span>${escapeHtml(payment.note)}</span>` : ""}
        </div>
      `).join("")}
    </div>
  `;
}

function getInvoiceReconciliationRows(invoices = state.invoices) {
  return invoices
    .flatMap((invoice) => (Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : []).map((payment) => ({
      id: payment.id || uid("payment"),
      invoiceId: invoice.id,
      invoiceNumber: invoice.invoiceNumber,
      clientLabel: getInvoiceClientLabel(invoice),
      paymentDate: payment.paymentDate || "",
      amount: Number(payment.amount || 0),
      note: String(payment.note || "").trim(),
      amountPaid: Number(invoice.amountPaid || 0),
      balanceDue: Number(invoice.balanceDue || 0),
      invoiceStatus: getInvoiceStatusLabel(invoice),
      invoiceRawStatus: invoice.status || "draft",
      lightDesigner: getInvoiceLightDesignerLabel(invoice),
      showNames: getInvoiceLinkedShowNames(invoice).join(", ")
    })))
    .sort((a, b) =>
      String(b.paymentDate || "").localeCompare(String(a.paymentDate || ""))
      || String(b.invoiceNumber || "").localeCompare(String(a.invoiceNumber || ""))
    );
}

function filterInvoiceReconciliationRows(rows = getInvoiceReconciliationRows()) {
  const query = String(state.ui.paymentReconSearchQuery || "").trim().toLowerCase();
  const year = state.ui.paymentReconYear || "all";
  const month = state.ui.paymentReconMonth || "all";
  const client = state.ui.paymentReconClient || "all";
  return rows.filter((row) => {
    const paymentYear = String(row.paymentDate || "").slice(0, 4);
    const paymentMonth = String(row.paymentDate || "").slice(0, 7);
    const haystack = [
      row.invoiceNumber,
      row.clientLabel,
      row.paymentDate,
      row.note,
      row.lightDesigner,
      row.showNames,
      row.invoiceStatus
    ].join(" ").toLowerCase();
    const matchesQuery = !query || haystack.includes(query);
    const matchesYear = year === "all" || paymentYear === year;
    const matchesMonth = month === "all" || paymentMonth === month;
    const matchesClient = client === "all" || row.clientLabel === client;
    return matchesQuery && matchesYear && matchesMonth && matchesClient;
  });
}

function getPaymentReconciliationYearOptions(rows = getInvoiceReconciliationRows()) {
  return [...new Set(rows.map((row) => String(row.paymentDate || "").slice(0, 4)).filter(Boolean))]
    .sort((a, b) => b.localeCompare(a));
}

function getPaymentReconciliationMonthOptions(rows = getInvoiceReconciliationRows(), selectedYear = "all") {
  return [...new Set(rows
    .filter((row) => selectedYear === "all" || String(row.paymentDate || "").slice(0, 4) === selectedYear)
    .map((row) => String(row.paymentDate || "").slice(0, 7))
    .filter(Boolean))]
    .sort((a, b) => b.localeCompare(a));
}

function makeDefaultInvoiceNumber() {
  const today = new Date();
  const prefix = `INV-${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, "0")}`;
  const existingCount = (state.invoices || []).filter((invoice) => String(invoice.invoiceNumber || "").startsWith(prefix)).length + 1;
  return `${prefix}-${String(existingCount).padStart(3, "0")}`;
}

function formatSyncStatus(show) {
  if (show.needsAdminCompletion) return "Needs Admin Completion";
  if (show.googleSyncStatus === "pending_push") return "Edited in PixelBug";
  if (show.googleSyncStatus === "updated_from_google") return "Updated from Google";
  if (show.googleSyncStatus === "sync_error") return "Sync Error";
  if (show.googleSyncStatus === "unlinked") return "Unlinked";
  if (show.googleSyncStatus === "synced") return "Synced";
  return "Imported";
}

function getArchiveYearOptions(shows) {
  return [...new Set(shows.map((show) => getShowStartDate(show).slice(0, 4)).filter(Boolean))].sort((a, b) => b.localeCompare(a));
}

function getArchiveMonthOptions(shows, selectedYear = "all") {
  const monthKeys = shows
    .filter((show) => selectedYear === "all" || getShowStartDate(show).slice(0, 4) === selectedYear)
    .map((show) => getShowStartDate(show).slice(0, 7))
    .filter(Boolean);
  return [...new Set(monthKeys)].sort((a, b) => b.localeCompare(a));
}

function filterArchivedGoogleShows(shows) {
  const selectedYear = state.ui.googleArchiveYear || "all";
  const selectedMonth = state.ui.googleArchiveMonth || "all";
  return shows.filter((show) => {
    const start = getShowStartDate(show);
    if (!start) return false;
    const year = start.slice(0, 4);
    const month = start.slice(0, 7);
    if (selectedYear !== "all" && year !== selectedYear) return false;
    if (selectedMonth !== "all" && month !== selectedMonth) return false;
    return true;
  });
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
      <select id="${selectId}" data-fallback-label="${allLabel}">
        <option value="all" ${state.ui.selectedCrewFilter === "all" ? "selected" : ""}>${allLabel}</option>
        ${includeUnassigned ? `<option value="unassigned" ${state.ui.selectedCrewFilter === "unassigned" ? "selected" : ""}>Unassigned</option>` : ""}
        ${crewUsers.map((crewUser) => `<option value="${crewUser.id}" ${state.ui.selectedCrewFilter === crewUser.id ? "selected" : ""}>${crewUser.name}</option>`).join("")}
      </select>
    </label>
  `;
}

function getSelectedCrewFilterLabel() {
  const selected = state.ui.selectedCrewFilter || "all";
  if (selected === "all") return "All Crew";
  if (selected === "unassigned") return "Unassigned";
  return getCrewUsers().find((user) => user.id === selected)?.name || "All Crew";
}

let customSelectOutsideHandlerAttached = false;

function closeAllCustomSelects(except = null) {
  document.querySelectorAll(".custom-select.open").forEach((wrapper) => {
    if (except && wrapper === except) return;
    wrapper.classList.remove("open");
    const trigger = wrapper.querySelector(".custom-select-trigger");
    if (trigger) trigger.setAttribute("aria-expanded", "false");
  });
}

function syncCustomSelect(select) {
  const wrapper = select.closest(".custom-select");
  if (!wrapper) return;
  const trigger = wrapper.querySelector(".custom-select-trigger");
  const selectedOption = [...select.options].find((option) => option.value === select.value) || select.options[0];
  const fallbackLabel = select.dataset.fallbackLabel || "";
  const displayLabel = selectedOption?.textContent?.trim() || fallbackLabel || getSelectedCrewFilterLabel();
  if (trigger) {
    trigger.innerHTML = `<span class="custom-select-value">${displayLabel}</span>`;
    trigger.title = displayLabel;
  }
  wrapper.querySelectorAll(".custom-select-option").forEach((optionButton) => {
    optionButton.classList.toggle("is-selected", optionButton.dataset.value === select.value);
  });
  const searchInput = wrapper.querySelector(".custom-select-search");
  if (searchInput) {
    searchInput.value = "";
    wrapper.querySelectorAll(".custom-select-option").forEach((optionButton) => {
      optionButton.classList.remove("hidden");
    });
  }
}

function enhanceCustomSelects(root = document) {
  if (!customSelectOutsideHandlerAttached) {
    document.addEventListener("click", (event) => {
      if (!event.target.closest(".custom-select")) {
        closeAllCustomSelects();
      }
    });
    customSelectOutsideHandlerAttached = true;
  }

  root.querySelectorAll("select").forEach((select) => {
    if (select.dataset.customSelectReady === "true") {
      syncCustomSelect(select);
      return;
    }

    select.dataset.customSelectReady = "true";
    select.classList.add("native-select-hidden");

    const wrapper = document.createElement("div");
    wrapper.className = "custom-select";

    const trigger = document.createElement("button");
    trigger.type = "button";
    trigger.className = "custom-select-trigger";
    trigger.setAttribute("aria-haspopup", "listbox");
    trigger.setAttribute("aria-expanded", "false");

    const menu = document.createElement("div");
    menu.className = "custom-select-menu";
    menu.setAttribute("role", "listbox");

    if (select.dataset.searchable === "true") {
      const search = document.createElement("input");
      search.type = "search";
      search.className = "custom-select-search";
      search.placeholder = select.dataset.searchPlaceholder || "Search...";
      search.autocomplete = "off";
      search.addEventListener("click", (event) => event.stopPropagation());
      search.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          event.stopPropagation();
        }
      });
      search.addEventListener("input", () => {
        const query = search.value.trim().toLowerCase();
        wrapper.querySelectorAll(".custom-select-option").forEach((optionButton) => {
          const matches = !query || optionButton.textContent.toLowerCase().includes(query);
          optionButton.classList.toggle("hidden", !matches);
        });
      });
      menu.appendChild(search);
    }

    [...select.options].forEach((option) => {
      const optionButton = document.createElement("button");
      optionButton.type = "button";
      optionButton.className = "custom-select-option";
      optionButton.dataset.value = option.value;
      optionButton.textContent = option.textContent;
      optionButton.addEventListener("click", () => {
        select.value = option.value;
        syncCustomSelect(select);
        wrapper.classList.remove("open");
        trigger.setAttribute("aria-expanded", "false");
        select.dispatchEvent(new Event("change", { bubbles: true }));
      });
      menu.appendChild(optionButton);
    });

    trigger.addEventListener("click", () => {
      const isOpen = wrapper.classList.contains("open");
      closeAllCustomSelects(wrapper);
      wrapper.classList.toggle("open", !isOpen);
      trigger.setAttribute("aria-expanded", String(!isOpen));
      if (!isOpen) {
        wrapper.querySelector(".custom-select-search")?.focus();
      }
    });

    select.parentNode.insertBefore(wrapper, select);
    wrapper.appendChild(select);
    wrapper.appendChild(trigger);
    wrapper.appendChild(menu);
    syncCustomSelect(select);
  });
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
        if (button.dataset.targetPanel !== state.ui.activeSidebarTab && !confirmDiscardDirtyForm("switch tabs")) {
          return;
        }
        clearDirtyForm();
        state.ui.activeSidebarTab = button.dataset.targetPanel;
        if (button.dataset.targetPanel === "calendarPanel") {
          state.view.mode = "month";
          state.ui.selectedCalendarShowId = null;
          state.ui.calendarReturnMode = "month";
        }
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
  const resolvedSelected = resolveCrewColor(selected);
  container.innerHTML = "";

  COLOR_OPTIONS.forEach((color) => {
    const resolvedColor = resolveCrewColor(color);
    if (taken.includes(resolvedColor) && resolvedSelected !== resolvedColor) {
      return;
    }
    const button = document.createElement("button");
    button.type = "button";
    button.className = `color-option ${resolvedSelected === resolvedColor ? "selected" : ""}`;
    button.style.background = resolvedColor;
    button.dataset.color = resolvedColor;
    button.title = resolvedColor;
    button.addEventListener("click", () => {
      container.querySelectorAll(".color-option").forEach((node) => node.classList.remove("selected"));
      button.classList.add("selected");
    });
    container.append(button);
  });
}

function renderSessionActions(user) {
  const node = document.getElementById("sessionActions");
  if (profileMenuOutsideHandler) {
    document.removeEventListener("click", profileMenuOutsideHandler);
    profileMenuOutsideHandler = null;
  }
  node.innerHTML = "";
  if (!user) return;

  const themePreference = state.ui.themePreference || "system";
  const themeButtonLabel = themePreference === "light"
    ? "Light Mode"
    : themePreference === "dark"
      ? "Dark Mode"
      : "Auto Theme";

  node.innerHTML = `
    <div class="profile-menu-wrap">
      <button type="button" class="ghost small ${state.ui.authPanelOpen ? "is-active" : ""}" id="profileMenuButton">Hey, ${user.name}</button>
      ${state.ui.authPanelOpen ? `
        <div class="profile-menu-panel">
          <div class="profile-menu-tabs">
            <button type="button" class="ghost small ${state.ui.authPanelMode === "profile" ? "is-active" : ""}" id="profilePanelButton">Info</button>
            <button type="button" class="ghost small ${state.ui.authPanelMode === "password" ? "is-active" : ""}" id="passwordPanelButton">Password</button>
            <button type="button" class="ghost small" id="themeCycleButton">Theme: ${themeButtonLabel}</button>
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
              <div class="meta-line">Role: ${getRoleLabel(user.role)}</div>
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

  document.getElementById("themeCycleButton")?.addEventListener("click", () => {
    const current = state.ui.themePreference || "system";
    const next = current === "light"
      ? "system"
      : current === "system"
        ? "dark"
        : "light";
    state.ui.themePreference = next;
    saveState(state);
    applyThemeFromState();
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

  if (state.ui.authPanelOpen) {
    profileMenuOutsideHandler = (event) => {
      if (!node.contains(event.target)) {
        state.ui.authPanelOpen = false;
        saveState(state);
        renderSessionActions(user);
      }
    };
    window.setTimeout(() => {
      if (profileMenuOutsideHandler) {
        document.addEventListener("click", profileMenuOutsideHandler);
      }
    }, 0);
  }
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
  title.textContent = isAdmin(user) ? "Admin Operations Board" : isAccounts(user) ? "Accounts Desk" : user.role === "viewer" ? "Calendar View" : "My Crew Schedule";
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

  if (state.ui.activeSidebarTab === "showsPanel") {
    singleView.innerHTML = `<section class="panel" id="showsPanel"></section>`;
    renderShowsPanel(user, shows, visibleShows);
    return;
  }

  if (state.ui.activeSidebarTab === "clientsPanel" && (isAdmin(user) || isAccounts(user))) {
    singleView.innerHTML = `<section class="panel" id="clientsPanel"></section>`;
    renderClientsPanel();
    return;
  }

  if (state.ui.activeSidebarTab === "invoicesPanel" && canAccessInvoices(user)) {
    singleView.innerHTML = `<section class="panel" id="invoicesPanel"></section>`;
    renderInvoicesPanel();
    return;
  }

  if (state.ui.activeSidebarTab === "documentsPanel" && canAccessInvoices(user)) {
    singleView.innerHTML = `<section class="panel" id="documentsPanel"></section>`;
    renderDocumentCenterPanel();
    return;
  }

  if (state.ui.activeSidebarTab === "googleEntriesPanel" && isAdmin(user)) {
    singleView.innerHTML = `<section class="panel" id="googleEntriesPanel"></section>`;
    renderGoogleEntriesPanel(user);
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
          <button type="button" class="ghost small calendar-nav-button ${state.view.mode === "month" ? "is-active" : ""}" data-view-mode="month">Month</button>
          <button type="button" class="ghost small calendar-nav-button ${state.view.mode === "week" ? "is-active" : ""}" data-view-mode="week">Week</button>
          <button type="button" class="ghost small calendar-nav-button ${state.view.mode === "day" ? "is-active" : ""}" data-view-mode="day">Day</button>
        </div>
      </div>
      <div class="calendar-center-group">
        <button type="button" class="ghost small calendar-nav-button" data-range-nav="prev">Previous</button>
        <button type="button" class="ghost small calendar-nav-button" data-range-nav="today">Today</button>
        <button type="button" class="ghost small calendar-nav-button" data-range-nav="next">Next</button>
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
  enhanceCustomSelects(panel);
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
  panel.querySelectorAll("[data-date-key]").forEach((cell) => {
    cell.addEventListener("click", (event) => {
      if (event.target.closest("[data-calendar-show-id], .calendar-more-toggle")) return;
      ensureUiState();
      state.ui.selectedCalendarShowId = null;
      state.ui.calendarReturnMode = "month";
      syncViewDateParts(parseDateKey(cell.dataset.dateKey));
      state.view.mode = "day";
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
  const manualCrewNames = show.assignments
    .filter((assignment) => !assignment.crewId && assignment.manualCrewName)
    .map((assignment) => assignment.manualCrewName);
  const palette = crewUsers.map((crewUser) => resolveCrewColor(crewUser.color)).filter(Boolean);
  const color = palette.length > 1
    ? `linear-gradient(135deg, ${palette.join(", ")})`
    : (palette[0] || "linear-gradient(135deg, #575e70, #8b93a5)");
  const visibleAssignees = user.role === "crew"
    ? crewUsers.filter((crewUser) => crewUser.id === user.id)
    : [...crewUsers, ...manualCrewNames.map((name) => ({ name }))];
  return { crewUsers, visibleAssignees, color };
}

function renderMonthRangeBar(show, user, startIndex, endIndex, laneIndex, clippedStart, clippedEnd) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);
  const span = (endIndex - startIndex) + 1;
  const left = `calc(${startIndex} * (((100% - (6 * var(--calendar-gap))) / 7) + var(--calendar-gap)) + 6px)`;
  const width = `calc(${span} * ((100% - (6 * var(--calendar-gap))) / 7) + (${Math.max(span - 1, 0)} * var(--calendar-gap)) - 12px)`;
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
  enhanceCustomSelects(panel);
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
  panel.querySelectorAll("[data-date-key]").forEach((cell) => {
    cell.addEventListener("click", (event) => {
      if (event.target.closest("[data-calendar-show-id]")) return;
      ensureUiState();
      state.ui.selectedCalendarShowId = null;
      state.ui.calendarReturnMode = "week";
      syncViewDateParts(parseDateKey(cell.dataset.dateKey));
      state.view.mode = "day";
      saveState(state);
      renderDashboard();
    });
  });
  wireCalendarToolbarControls(panel);
}

function renderWeekRangeBar(show, user, startIndex, endIndex, laneIndex, clippedStart, clippedEnd) {
  const { visibleAssignees, color } = getShowDisplayMeta(show, user);
  const span = (endIndex - startIndex) + 1;
  const left = `calc(${startIndex} * (((100% - (6 * var(--calendar-gap))) / 7) + var(--calendar-gap)) + 6px)`;
  const width = `calc(${span} * ((100% - (6 * var(--calendar-gap))) / 7) + (${Math.max(span - 1, 0)} * var(--calendar-gap)) - 12px)`;
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
      <div class="lane-day-header single day-view-header ${focusKey === dateKey(new Date()) ? "today" : ""}">
        <div>
          ${focusDate.toLocaleDateString("en-IN", { weekday: "long" })}
          <strong>${formatWeekdayDate(focusDate)}</strong>
        </div>
        ${isAdmin(user) ? `<button type="button" class="secondary small calendar-add-show-button" data-create-show-date="${focusKey}">+ Add Show</button>` : ""}
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
  enhanceCustomSelects(panel);
  panel.querySelector("#closeCalendarShowPanel")?.addEventListener("click", () => {
    ensureUiState();
    state.ui.selectedCalendarShowId = null;
    state.view.mode = state.ui.calendarReturnMode || "month";
    state.ui.calendarReturnMode = "month";
    saveState(state);
    renderDashboard();
  });
  panel.querySelectorAll("[data-edit-show]").forEach((button) => {
    button.addEventListener("click", () => fillShowForm(button.dataset.editShow));
  });
  panel.querySelector("[data-create-show-date]")?.addEventListener("click", (event) => {
    startShowDraftForDate(event.currentTarget.dataset.createShowDate);
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
  const manualCrewNames = show.assignments
    .filter((assignment) => !assignment.crewId && assignment.manualCrewName)
    .map((assignment) => assignment.manualCrewName);
  const palette = crewUsers.map((crewUser) => crewUser.color).filter(Boolean);
  const color = palette.length > 1 ? `linear-gradient(135deg, ${palette.join(", ")})` : (palette[0] || "#264653");
  const visibleAssignees = user.role === "crew" ? crewUsers.filter((crewUser) => crewUser.id === user.id) : [...crewUsers, ...manualCrewNames.map((name) => ({ name }))];
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
            <div><span class="legend-swatch" style="background:${resolveCrewColor(crewUser.color)}"></span><strong>${crewUser.name}</strong></div>
            <div class="meta">${crewUser.phone}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderShowsList(user, shows, sourceShows = shows) {
  const panel = document.getElementById("showsListPanel") || document.getElementById("showsPanel");
  const showSearchQuery = String(state.ui.showSearchQuery || "").trim().toLowerCase();
  const todayKey = dateKey(new Date());
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
    const searchableValues = [
      show.showName,
      show.client,
      show.location,
      show.venue,
      ...show.assignments.map((assignment) => getAssignmentCrewName(assignment))
    ];
    const searchMatch = !showSearchQuery || searchableValues.some((value) => String(value || "").toLowerCase().includes(showSearchQuery));
    return yearMatch && monthMatch && searchMatch;
  });
  const currentShows = filteredShows.filter((show) => getShowStartDate(show) <= todayKey && getShowEndDate(show) >= todayKey);
  const upcomingShows = filteredShows.filter((show) => getShowStartDate(show) > todayKey);
  const pastShows = filteredShows.filter((show) => getShowEndDate(show) < todayKey);
  const timelineMode = state.ui.showTimelineMode || "active";
  const visibleShows = timelineMode === "past"
    ? sortShows(pastShows, state.ui.showSortMode)
    : sortShows([...currentShows, ...upcomingShows], state.ui.showSortMode);
  const activeMonthLabel = state.ui.activeShowMonth !== "all"
    ? monthGroupLabel(`${state.ui.activeShowMonth}-01`)
    : state.ui.selectedShowYear !== "all"
      ? `All Months in ${state.ui.selectedShowYear}`
      : "All Months";
  const selectedDraftShowIds = new Set(state.ui.invoiceDraftShowIds || []);
  const showPagination = getPaginationSlice(visibleShows, "showsPage", "showsPageSize");
  const pagedCurrentShows = showPagination.items.filter((show) => getShowStartDate(show) <= todayKey && getShowEndDate(show) >= todayKey);
  const pagedUpcomingShows = showPagination.items.filter((show) => getShowStartDate(show) > todayKey);
  const pagedPastShows = showPagination.items.filter((show) => getShowEndDate(show) < todayKey);

  panel.innerHTML = `
    <div class="stack">
      <div>
        <h3>${isAdmin(user) ? "All Show Entries" : "Visible Show Entries"}</h3>
      </div>
      ${sourceShows.length ? `
        <div class="shows-toolbar">
          <div class="shows-toolbar-top">
            <label class="sort-control invoice-search-control">
              <span>Search</span>
              <input type="search" id="showSearchInput" placeholder="Show, client, location, crew" value="${escapeHtml(state.ui.showSearchQuery || "")}" autocomplete="off">
            </label>
            <button type="button" class="secondary search-submit-button" id="applyShowSearchButton">Search</button>
          </div>
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
          </div>
          ${isAdmin(user) ? `
            <div class="shows-selection-toolbar">
              <button type="button" class="secondary" id="createInvoiceFromShowsButton">Create Invoice From Selected</button>
              <button type="button" class="ghost" id="clearInvoiceShowSelectionButton">Clear Selection</button>
            </div>
          ` : ""}
        </div>
        <section class="month-group">
          <div class="invoice-subtabs" role="tablist" aria-label="Show timeline">
            <button type="button" class="${timelineMode === "active" ? "is-active" : ""}" data-show-timeline="active">Current & Upcoming</button>
            <button type="button" class="${timelineMode === "past" ? "is-active" : ""}" data-show-timeline="past">Past Shows</button>
          </div>
          <header class="month-group-header">
            <h4>${timelineMode === "past" ? `Past Shows · ${activeMonthLabel}` : `Current & Upcoming · ${activeMonthLabel}`}</h4>
            <span class="pill">${visibleShows.length} ${visibleShows.length === 1 ? "show" : "shows"}</span>
          </header>
          ${timelineMode === "past" ? `
            <div class="show-list">
              ${pagedPastShows.length ? pagedPastShows.map((show) => renderShowCard(show, user, selectedDraftShowIds.has(show.id))).join("") : "<p>No past shows match the current filters.</p>"}
            </div>
          ` : `
            <div class="stack tight">
              <div class="month-group-header">
                <h4>Current Shows</h4>
                <span class="pill">${currentShows.length}</span>
              </div>
              <div class="show-list">
                ${pagedCurrentShows.length ? pagedCurrentShows.map((show) => renderShowCard(show, user, selectedDraftShowIds.has(show.id))).join("") : "<p>No current shows match the current filters.</p>"}
              </div>
            </div>
            <div class="stack tight">
              <div class="month-group-header">
                <h4>Upcoming Shows</h4>
                <span class="pill">${upcomingShows.length}</span>
              </div>
              <div class="show-list">
                ${pagedUpcomingShows.length ? pagedUpcomingShows.map((show) => renderShowCard(show, user, selectedDraftShowIds.has(show.id))).join("") : "<p>No upcoming shows match the current filters.</p>"}
              </div>
            </div>
          `}
          ${renderPaginationControls("shows", showPagination, "shows")}
        </section>
      ` : "<p>No shows available in your current view.</p>"}
    </div>
  `;

  enhanceCustomSelects(panel);

  const yearFilterSelect = document.getElementById("showYearFilter");
  panel.querySelectorAll("[data-show-timeline]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.showTimelineMode = button.dataset.showTimeline || "active";
      state.ui.showsPage = 1;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  });
  const showSearchInput = document.getElementById("showSearchInput");
  const applyShowSearch = () => {
    state.ui.showSearchQuery = showSearchInput?.value || "";
    state.ui.showsPage = 1;
    saveState(state);
    renderShowsList(user, shows, sourceShows);
  };
  showSearchInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    applyShowSearch();
  });
  document.getElementById("applyShowSearchButton")?.addEventListener("click", applyShowSearch);

  if (yearFilterSelect) {
    yearFilterSelect.addEventListener("change", () => {
      state.ui.selectedShowYear = yearFilterSelect.value;
      state.ui.activeShowMonth = "all";
      state.ui.showsPage = 1;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const monthFilterSelect = document.getElementById("showMonthFilter");
  if (monthFilterSelect) {
    monthFilterSelect.addEventListener("change", () => {
      state.ui.activeShowMonth = monthFilterSelect.value;
      state.ui.showsPage = 1;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const sortSelect = document.getElementById("showSortMode");
  if (sortSelect) {
    sortSelect.addEventListener("change", () => {
      state.ui.showSortMode = sortSelect.value;
      state.ui.showsPage = 1;
      saveState(state);
      renderShowsList(user, shows, sourceShows);
    });
  }

  const crewFilterSelect = document.getElementById("showsCrewFilter");
  if (crewFilterSelect) {
    crewFilterSelect.addEventListener("change", () => {
      state.ui.selectedCrewFilter = crewFilterSelect.value;
      state.ui.showsPage = 1;
      saveState(state);
      renderDashboard();
    });
  }

  wirePaginationControls(panel, "shows", "showsPage", "showsPageSize", () => renderShowsList(user, shows, sourceShows));

  const createInvoiceButton = document.getElementById("createInvoiceFromShowsButton");
  const clearInvoiceSelectionButton = document.getElementById("clearInvoiceShowSelectionButton");
  if (createInvoiceButton && isAdmin(user)) {
    const syncCreateButtonState = () => {
      const selectedCount = panel.querySelectorAll('[data-select-show-for-invoice]:checked').length;
      createInvoiceButton.disabled = selectedCount === 0;
      if (clearInvoiceSelectionButton) {
        clearInvoiceSelectionButton.disabled = selectedCount === 0;
      }
      createInvoiceButton.textContent = selectedCount ? `Create Invoice From Selected (${selectedCount})` : "Create Invoice From Selected";
    };

    panel.querySelectorAll('[data-select-show-for-invoice]').forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        const { showId } = checkbox.dataset;
        if (!showId) return;
        const next = new Set(state.ui.invoiceDraftShowIds || []);
        if (checkbox.checked) {
          next.add(showId);
        } else {
          next.delete(showId);
        }
        state.ui.invoiceDraftShowIds = [...next];
        saveState(state);
        syncCreateButtonState();
      });
    });

    createInvoiceButton.addEventListener("click", () => {
      const selectedIds = [...panel.querySelectorAll('[data-select-show-for-invoice]:checked')].map((checkbox) => checkbox.dataset.showId).filter(Boolean);
      if (!selectedIds.length) return;
      state.ui.invoiceDraftShowIds = selectedIds;
      state.ui.editingInvoiceId = null;
      state.ui.activeSidebarTab = "invoicesPanel";
      saveState(state);
      renderSidebarTabs();
      renderDashboard();
      showToast("Invoice draft ready from selected shows.");
    });

    clearInvoiceSelectionButton?.addEventListener("click", () => {
      panel.querySelectorAll('[data-select-show-for-invoice]:checked').forEach((checkbox) => {
        checkbox.checked = false;
      });
      state.ui.invoiceDraftShowIds = [];
      saveState(state);
      syncCreateButtonState();
      showToast("Selection cleared.");
    });

    syncCreateButtonState();
  }

  if (isAdmin(user) || isAccounts(user)) {
    panel.querySelectorAll("[data-edit-show]").forEach((button) => {
      button.addEventListener("click", () => fillShowForm(button.dataset.editShow));
    });
  }

  if (isAdmin(user)) {
    panel.querySelectorAll("[data-duplicate-show]").forEach((button) => {
      button.addEventListener("click", () => startShowDraftFromExistingShow(button.dataset.duplicateShow));
    });
  }
}

function renderShowsPanel(user, shows, sourceShows = shows) {
  const panel = document.getElementById("showsPanel");
  const isEditing = Boolean(state.ui.editingShowId);
  const canManageShows = isAdmin(user);
  const activeShowSubtab = canManageShows ? (isEditing ? "create" : state.ui.showSubtab || "list") : "list";

  if (!canManageShows) {
    renderShowsList(user, shows, sourceShows);
    return;
  }

  panel.innerHTML = `
    <div class="stack">
      <div class="form-header">
        <div>
          <h3>Shows</h3>
          <p class="muted-note">Create new show entries and manage the complete show register from one place.</p>
        </div>
      </div>
      <div class="invoice-subtabs" role="tablist" aria-label="Show sections">
        <button type="button" class="${activeShowSubtab === "create" ? "is-active" : ""}" data-show-subtab="create">${getDirtyTabLabel("Create Show", "show")}</button>
        <button type="button" class="${activeShowSubtab === "list" ? "is-active" : ""}" data-show-subtab="list">All Shows</button>
      </div>
      <section id="showFormPanel" class="${activeShowSubtab === "create" ? "" : "hidden"}"></section>
      <section id="showsListPanel" class="${activeShowSubtab === "list" ? "" : "hidden"}"></section>
    </div>
  `;

  renderShowForm();
  renderShowsList(user, shows, sourceShows);

  panel.querySelectorAll("[data-show-subtab]").forEach((button) => {
    button.addEventListener("click", () => {
      if ((button.dataset.showSubtab || "list") !== activeShowSubtab && !confirmDiscardDirtyForm("switch show sections")) {
        return;
      }
      clearDirtyForm();
      state.ui.showSubtab = button.dataset.showSubtab || "list";
      if (state.ui.showSubtab === "list") {
        resetEditingState();
      }
      saveState(state);
      renderDashboard();
    });
  });
}

function renderShowCard(show, user, selectedForInvoice = false) {
  const assignments = show.assignments.map((assignment) => {
    const crewUser = getUserById(assignment.crewId);
    const crewName = getAssignmentCrewName(assignment);
    if (!crewName) return "";
    const lightDesigner = assignment.lightDesignerId ? getUserById(assignment.lightDesignerId) : null;
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
            <strong>${escapeHtml(crewName)}</strong>
          </div>
          <span class="pill" style="background:${crewUser ? resolveCrewColor(crewUser.color) : "#8b93a5"}; color:white;">${crewUser ? "Crew" : "Manual Crew"}</span>
        </header>
        <details class="assignment-details">
          <summary>
            <span class="more-label">More..</span>
            <span class="less-label">Less..</span>
          </summary>
          <div class="assignment-details-body">
            <div class="meta">Light Designer: ${lightDesigner?.name || "-"}</div>
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
        ${isAdmin(user) ? `
          <div class="toolbar">
            <label class="show-select-chip">
              <input type="checkbox" data-select-show-for-invoice data-show-id="${show.id}" ${selectedForInvoice ? "checked" : ""}>
              <span>Bill</span>
            </label>
            <button type="button" class="ghost small" data-duplicate-show="${show.id}">Duplicate</button>
            <button type="button" class="secondary small" data-edit-show="${show.id}">Edit</button>
          </div>
        ` : ""}
      </header>
      <div class="show-banner">
        <span class="show-banner-item">${show.location || "Location TBD"}</span>
        <span class="show-banner-item">${escapeHtml(getClientDisplayValue(show.clientId, show.client) || "Client TBD")}</span>
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

function formatInvoiceDate(dateStr) {
  return dateStr ? formatDate(dateStr) : "-";
}

function parseInvoiceShowIds(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  const normalized = String(value || "").trim();
  if (!normalized) return [];
  if (normalized.startsWith("[") && normalized.endsWith("]")) {
    try {
      return JSON.parse(normalized).map((item) => String(item || "").trim()).filter(Boolean);
    } catch (error) {
      return [];
    }
  }
  return normalized.split(",").map((item) => item.trim()).filter(Boolean);
}

function serializeInvoiceShowIds(showIds = []) {
  return [...new Set(showIds.map((showId) => String(showId || "").trim()).filter(Boolean))].join(",");
}

function getInvoiceLineShowLabel(showId) {
  return getInvoiceLineShowLabels(showId).join(", ");
}

function getInvoiceLineShowLabels(showId) {
  return parseInvoiceShowIds(showId).map((id) => {
    const show = state.shows.find((item) => item.id === id);
    if (!show) return "";
    const location = String(show.location || show.venue || "").trim();
    return `${show.showName} · ${formatDateRange(getShowStartDate(show), getShowEndDate(show))}${location ? ` · ${location}` : ""}`;
  }).filter(Boolean);
}

function defaultInvoiceLineDescription(show) {
  if (!show) return "";
  const dateLabel = formatDateRange(getShowStartDate(show), getShowEndDate(show));
  return `${show.showName}${dateLabel !== "-" ? ` (${dateLabel})` : ""}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatMultilineHtml(value, emptyText = "") {
  const text = String(value || "").trim();
  if (!text) {
    return emptyText ? `<span class="invoice-print-subtle">${escapeHtml(emptyText)}</span>` : "";
  }
  return escapeHtml(text).replace(/\n/g, "<br>");
}

function numberToIndianWords(value) {
  const amount = Math.round(Number(value || 0));
  if (!amount) return "Zero rupees only";
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  const underHundred = (number) => number < 20 ? ones[number] : `${tens[Math.floor(number / 10)]}${number % 10 ? ` ${ones[number % 10]}` : ""}`;
  const underThousand = (number) => {
    const hundred = Math.floor(number / 100);
    const rest = number % 100;
    return `${hundred ? `${ones[hundred]} Hundred${rest ? " " : ""}` : ""}${rest ? underHundred(rest) : ""}`;
  };
  const parts = [];
  const crore = Math.floor(amount / 10000000);
  const lakh = Math.floor((amount % 10000000) / 100000);
  const thousand = Math.floor((amount % 100000) / 1000);
  const rest = amount % 1000;
  if (crore) parts.push(`${underThousand(crore)} Crore`);
  if (lakh) parts.push(`${underThousand(lakh)} Lakh`);
  if (thousand) parts.push(`${underThousand(thousand)} Thousand`);
  if (rest) parts.push(underThousand(rest));
  return `${parts.join(" ")} rupees only`;
}

function getInvoiceCopyLabels(copyCount = 1) {
  const count = Math.max(1, Math.min(5, Number(copyCount || 1)));
  const labels = {
    1: ["Original Copy"],
    2: ["Supplier Copy", "Recipient Copy"],
    3: ["Supplier Copy", "Transporter Copy", "Recipient Copy"],
    4: ["Original Copy", "Duplicate Copy", "Triplicate Copy", "Additional Copy"],
    5: ["Original Copy", "Duplicate Copy", "Triplicate Copy", "Additional Copy 1", "Additional Copy 2"]
  };
  return labels[count] || labels[1];
}

function getInvoiceHsnSacSummaryRows(lineItems = [], placeOfSupply = "") {
  const summaryMap = new Map();
  const intraState = isMaharashtraSupply(placeOfSupply);
  lineItems.forEach((item) => {
    const baseAmount = Number(item.quantity || 1) * Number(item.unitRate || 0);
    const discountAmount = getDiscountAmount(item.discount || "", baseAmount);
    const taxableAmount = Math.max(0, baseAmount - discountAmount);
    const key = String(item.sac || "").trim() || "-";
    const current = summaryMap.get(key) || { sac: key, taxableAmount: 0 };
    current.taxableAmount += taxableAmount;
    summaryMap.set(key, current);
  });
  return [...summaryMap.values()].map((row) => {
    const taxableAmount = Math.round(row.taxableAmount * 100) / 100;
    const sgstAmount = intraState ? Math.round(taxableAmount * 0.09 * 100) / 100 : 0;
    const cgstAmount = intraState ? Math.round(taxableAmount * 0.09 * 100) / 100 : 0;
    const igstAmount = intraState ? 0 : Math.round(taxableAmount * 0.18 * 100) / 100;
    const gstAmount = Math.round((sgstAmount + cgstAmount + igstAmount) * 100) / 100;
    return {
      sac: row.sac,
      taxableAmount,
      gstRateLabel: intraState ? "SGST 9% + CGST 9%" : "IGST 18%",
      gstAmount,
      totalAmount: Math.round((taxableAmount + gstAmount) * 100) / 100
    };
  });
}

function getSingleInvoiceDocumentMarkup(invoice, copyLabel = "Original Copy") {
  const lightDesigner = getInvoiceLightDesignerLabel(invoice) || "-";
  const details = normalizeInvoiceDetails(invoice.details);
  const invoiceClient = getClientById(invoice.clientId) || getClientByName(invoice.clientName);
  const effectiveDueDate = invoice.dueDate || getDueDateFromTerms(invoice.issueDate, details.paymentTerms);
  const effectivePlaceOfSupply = details.placeOfSupply || invoiceClient?.state || "";
  const effectiveClientGstin = details.clientGstin || invoiceClient?.gstin || "";
  const companyProfile = getFixedInvoiceCompanyProfile();
  const gstBreakup = getInvoiceCalculationFromValues(invoice.lineItems || [], effectivePlaceOfSupply);
  const hsnSacSummaryRows = getInvoiceHsnSacSummaryRows(invoice.lineItems || [], effectivePlaceOfSupply);
  const invoiceAmountPaid = Number(invoice.amountPaid || 0);
  const invoiceBalanceDue = Math.max(0, gstBreakup.totalAmount - invoiceAmountPaid);
  const bankRows = [
    { label: "Account Name", value: companyProfile.bankAccountName },
    { label: "Bank", value: companyProfile.bankName },
    { label: "Account Number", value: companyProfile.bankAccountNumber },
    { label: "Account Type", value: companyProfile.bankAccountType },
    { label: "Branch Name", value: companyProfile.bankBranchName },
    { label: "IFSC", value: companyProfile.bankIfsc }
  ].filter((item) => item.value);
  const notesMarkup = String(invoice.notes || "").trim()
    ? `<p>${escapeHtml(invoice.notes)}</p>`
    : `<div class="invoice-print-empty-space"></div>`;
  const lineItemsMarkup = (invoice.lineItems || []).map((item, index) => {
    const grossLineAmount = Number(item.quantity || 1) * Number(item.unitRate || 0);
    const discountAmount = getDiscountAmount(item.discount || "", grossLineAmount);
    const netLineAmount = Math.max(0, grossLineAmount - discountAmount);
    const linkedShowLines = getInvoiceLineShowLabels(item.showId);
    return `
      <tr>
        <td>${index + 1}</td>
        <td>
          <strong>${escapeHtml(item.description)}</strong>
          ${item.customDetails ? `<div class="invoice-print-subtle">${escapeHtml(item.customDetails)}</div>` : ""}
          ${linkedShowLines.map((label) => `<div class="invoice-print-subtle">${escapeHtml(label)}</div>`).join("")}
        </td>
        <td>${escapeHtml(item.sac || "")}</td>
        <td>${Number(item.quantity || 0)}</td>
        <td>${escapeHtml(formatCurrency(item.unitRate))}</td>
        <td>${escapeHtml(formatCurrency(netLineAmount))}</td>
      </tr>
    `;
  }).join("");

  return `
    <div class="invoice-print-page">
      <header class="invoice-print-header">
        <div class="invoice-print-brand-block">
          <div class="invoice-print-brand-row">
            <img src="PIXELBUG.JPG" alt="PixelBug logo" class="invoice-print-logo">
            <div class="invoice-print-brand-copy">
              <div class="invoice-print-brand-title">${escapeHtml(companyProfile.name)}</div>
              <div class="invoice-print-company-line">${formatMultilineHtml(companyProfile.address)}</div>
              ${companyProfile.gstin ? `<div class="invoice-print-company-line">GSTIN: ${escapeHtml(companyProfile.gstin)}</div>` : ""}
              ${companyProfile.phone ? `<div class="invoice-print-company-line">${escapeHtml(companyProfile.phone)}</div>` : ""}
              ${companyProfile.email ? `<div class="invoice-print-company-line">${escapeHtml(companyProfile.email)}</div>` : ""}
              ${companyProfile.website ? `<div class="invoice-print-company-line">${escapeHtml(companyProfile.website)}</div>` : ""}
            </div>
            <div class="invoice-print-copy-block">
              <div class="invoice-print-document-title">Tax Invoice</div>
              <div class="invoice-print-copy-label">${escapeHtml(copyLabel)}</div>
            </div>
          </div>
        </div>
      </header>
      <section class="invoice-print-section invoice-print-grid invoice-print-summary-grid">
        <div class="invoice-print-card">
          <h2>Invoice Details</h2>
          <div class="invoice-print-detail-list">
            <div><span>Invoice No</span><strong>${escapeHtml(invoice.invoiceNumber)}</strong></div>
            <div><span>Issue Date</span><strong>${escapeHtml(formatInvoiceDate(invoice.issueDate))}</strong></div>
            <div><span>Payment Terms</span><strong>${escapeHtml(details.paymentTerms || "Net 15")}</strong></div>
            <div><span>Due Date</span><strong>${escapeHtml(formatInvoiceDate(effectiveDueDate))}</strong></div>
          </div>
        </div>
        <div class="invoice-print-card">
          <h2>Additional Details</h2>
          <div class="invoice-print-detail-list">
            <div><span>Status</span><strong>${escapeHtml(getInvoiceStatusLabel(invoice))}</strong></div>
            <div><span>Place of Supply</span><strong>${escapeHtml(effectivePlaceOfSupply || "-")}</strong></div>
            <div><span>Light Designer</span><strong>${escapeHtml(lightDesigner)}</strong></div>
          </div>
        </div>
      </section>
      <div class="invoice-print-billto-rule"></div>
      <section class="invoice-print-section invoice-print-billto">
        <h2>Bill To</h2>
        <p>${escapeHtml(invoice.clientName)}</p>
        <p>${formatMultilineHtml(details.clientBillingAddress, "Add client billing address in invoice details.")}</p>
        ${effectiveClientGstin ? `<p>GSTIN: ${escapeHtml(effectiveClientGstin)}</p>` : ""}
        <div class="invoice-print-billto-divider"></div>
        <p><strong>Notes:</strong> ${String(invoice.notes || "").trim() ? escapeHtml(invoice.notes) : ""}</p>
      </section>
      <div class="invoice-print-notes-rule"></div>
      <section class="invoice-print-section">
        <table class="invoice-print-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Particulars</th>
              <th>SAC</th>
              <th>Qty</th>
              <th>Rate</th>
              <th>Amount</th>
            </tr>
          </thead>
          <tbody>${lineItemsMarkup}</tbody>
        </table>
      </section>
      <section class="invoice-print-section invoice-print-totals">
        <div class="invoice-print-amount-words">
          <div class="invoice-print-amount-line">
            <span>Total In Words</span>
            <strong>${escapeHtml(numberToIndianWords(gstBreakup.totalAmount))}</strong>
          </div>
          ${bankRows.length ? `
            <div class="invoice-print-bank-block">
              <p>In case of online transfer, bank details are as follows</p>
              ${bankRows.map((item) => `<div><span>${escapeHtml(item.label)}:</span> <strong>${escapeHtml(item.value)}</strong></div>`).join("")}
            </div>
          ` : `
            <div class="invoice-print-bank-block">
              <p>In case of online transfer, bank details are as follows</p>
              <p class="invoice-print-subtle">Bank details not configured yet.</p>
            </div>
          `}
        </div>
        <div><span>Subtotal</span><strong>${escapeHtml(formatCurrency(gstBreakup.grossSubtotal))}</strong></div>
        <div><span>Discount</span><strong>${gstBreakup.discountAmount ? escapeHtml(formatCurrency(gstBreakup.discountAmount)) : "-"}</strong></div>
        <div><span>SGST (9%)</span><strong>${gstBreakup.sgstAmount ? escapeHtml(formatCurrency(gstBreakup.sgstAmount)) : "-"}</strong></div>
        <div><span>CGST (9%)</span><strong>${gstBreakup.cgstAmount ? escapeHtml(formatCurrency(gstBreakup.cgstAmount)) : "-"}</strong></div>
        <div><span>IGST (18%)</span><strong>${gstBreakup.igstAmount ? escapeHtml(formatCurrency(gstBreakup.igstAmount)) : "-"}</strong></div>
        <div class="invoice-print-total-row"><span>Total</span><strong>${escapeHtml(formatCurrency(gstBreakup.totalAmount))}</strong></div>
        <div><span>Amount Paid</span><strong>${escapeHtml(formatCurrency(invoiceAmountPaid))}</strong></div>
        <div class="invoice-print-balance"><span>Balance Due</span><strong>${escapeHtml(formatCurrency(invoiceBalanceDue))}</strong></div>
        <div class="invoice-print-balance-rule"></div>
        <div class="invoice-print-signature">
          <div class="invoice-print-signature-space">
            ${companyProfile.signatureImage ? `<img src="${escapeHtml(companyProfile.signatureImage)}" alt="Authorised signature" class="invoice-print-signature-image">` : ""}
          </div>
          ${companyProfile.signatureHolder ? `<strong>${escapeHtml(companyProfile.signatureHolder)}</strong>` : ""}
          <span>Authorised Signatory</span>
        </div>
      </section>
      ${hsnSacSummaryRows.length ? `
        <section class="invoice-print-section invoice-print-hsn-summary">
          <h2>HSN/SAC Summary</h2>
          <table class="invoice-print-hsn-summary-table">
            <thead>
              <tr>
                <th>HSN/SAC</th>
                <th>Taxable Amount</th>
                <th>GST Rate</th>
                <th>GST Amount</th>
                <th>Total Amount</th>
              </tr>
            </thead>
            <tbody>
              ${hsnSacSummaryRows.map((row) => `
                <tr>
                  <td>${escapeHtml(row.sac)}</td>
                  <td>${escapeHtml(formatCurrency(row.taxableAmount))}</td>
                  <td>${escapeHtml(row.gstRateLabel)}</td>
                  <td>${escapeHtml(formatCurrency(row.gstAmount))}</td>
                  <td>${escapeHtml(formatCurrency(row.totalAmount))}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </section>
      ` : ""}
      <footer class="invoice-print-footer">
        <div class="invoice-print-thank-you">Thank you for your business.</div>
        <div class="invoice-print-payment-note">Please clear the dues within 15 working days.</div>
        ${companyProfile.footerNote ? `<div class="invoice-print-payment-note">${escapeHtml(companyProfile.footerNote)}</div>` : ""}
      </footer>
      <div class="invoice-print-page-number"></div>
    </div>
  `;
}

function getInvoiceDocumentMarkup(invoice, options = {}) {
  const copyLabels = Array.isArray(options.copyLabels) && options.copyLabels.length
    ? options.copyLabels
    : getInvoiceCopyLabels(options.copyCount || 1);
  return copyLabels.map((copyLabel) => getSingleInvoiceDocumentMarkup(invoice, copyLabel)).join("");
}

function getInvoicePrintStyles() {
  return `
    @page {
      size: A4;
      margin: 0;
      @bottom-right {
        content: "Page " counter(page);
        color: #667789;
        font-size: 9px;
      }
    }
    * { box-sizing: border-box; }
    body { margin: 0; background: #eef2f7; color: #17212b; font-family: "IBM Plex Sans", sans-serif; }
    .invoice-print-page { width: 210mm; min-height: 297mm; margin: 0 auto; background: #fff; padding: 10mm; font-size: 11px; box-shadow: none; }
    .invoice-print-header { display: block; margin-bottom: 16px; border-bottom: 2px solid #d8e0ea; padding-bottom: 12px; }
    .invoice-print-brand-block { flex: 1 1 auto; }
    .invoice-print-brand-row { display: flex; align-items: flex-start; gap: 12px; }
    .invoice-print-logo { width: 70px; height: 70px; object-fit: contain; border-radius: 12px; }
    .invoice-print-brand-copy { flex: 1 1 auto; }
    .invoice-print-brand-title { margin: 0 0 5px; font-family: "Space Grotesk", sans-serif; font-size: 24px; font-weight: 700; line-height: 1.05; }
    .invoice-print-company-line { margin: 0 0 2px; color: #4a5b6d; font-size: 10px; line-height: 1.24; }
    .invoice-print-copy-block { margin-left: auto; text-align: right; }
    .invoice-print-document-title { font-family: "Space Grotesk", sans-serif; font-size: 24px; font-weight: 700; line-height: 1.05; text-align: right; white-space: nowrap; }
    .invoice-print-copy-label { margin-top: 5px; color: #5c6b7a; font-size: 10px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; text-align: right; }
    .invoice-print-subtle { color: #667789; font-size: 10px; line-height: 1.3; }
    .invoice-print-meta { min-width: 260px; display: grid; gap: 12px; padding: 18px; border: 1px solid #dfe6ef; border-radius: 18px; background: #f8fafc; }
    .invoice-print-meta div { display: flex; justify-content: space-between; gap: 18px; }
    .invoice-print-meta span { color: #667789; font-size: 13px; }
    .invoice-print-section { margin-bottom: 12px; }
    .invoice-print-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .invoice-print-summary-grid { align-items: start; }
    .invoice-print-card { padding: 11px; border: 1px solid #dfe6ef; border-radius: 10px; background: #fbfcfe; min-height: 100%; }
    .invoice-print-grid h2 { margin: 0 0 5px; font-size: 10px; text-transform: uppercase; letter-spacing: .08em; color: #5c6b7a; }
    .invoice-print-grid p { margin: 0 0 4px; line-height: 1.3; }
    .invoice-print-empty-space { min-height: 18px; }
    .invoice-print-billto-rule, .invoice-print-notes-rule { margin: 2px 0 9px; border-top: 2px solid #17212b; }
    .invoice-print-billto { margin-bottom: 8px; }
    .invoice-print-billto h2 { margin: 0 0 4px; color: #5c6b7a; font-size: 10px; letter-spacing: .08em; text-transform: uppercase; }
    .invoice-print-billto p { margin: 0 0 3px; line-height: 1.28; }
    .invoice-print-billto-divider { margin: 6px 0 5px; border-top: 1px solid #dfe6ef; }
    .invoice-print-detail-list { display: grid; gap: 5px; }
    .invoice-print-detail-list div { display: flex; justify-content: space-between; gap: 12px; border-bottom: 1px solid #dfe6ef; padding-bottom: 4px; }
    .invoice-print-detail-list span { color: #667789; font-size: 10px; }
    .invoice-print-table { width: 100%; border-collapse: collapse; table-layout: auto; border: 0; }
    .invoice-print-table th, .invoice-print-table td { padding: 7px 8px; border: 0; border-bottom: 1px solid #dfe6ef; text-align: left; vertical-align: top; }
    .invoice-print-table th { color: #5c6b7a; font-size: 9px; text-transform: uppercase; letter-spacing: .08em; background: #f8fafc; }
    .invoice-print-hsn-summary { margin-top: 6px; padding-top: 2px; }
    .invoice-print-hsn-summary h2 { margin: 0 0 6px; color: #5c6b7a; font-size: 10px; letter-spacing: .08em; text-transform: uppercase; }
    .invoice-print-hsn-summary-table { width: 100%; border-collapse: collapse; table-layout: auto; }
    .invoice-print-hsn-summary-table th, .invoice-print-hsn-summary-table td { padding: 5px 8px; border-bottom: 1px solid #dfe6ef; text-align: left; vertical-align: top; }
    .invoice-print-hsn-summary-table th { color: #5c6b7a; font-size: 9px; text-transform: uppercase; letter-spacing: .08em; background: #f8fafc; }
    .invoice-print-totals { margin-left: 0; display: grid; grid-template-columns: 1fr minmax(250px, 300px); column-gap: 12px; row-gap: 0; }
    .invoice-print-totals > div:not(.invoice-print-amount-words) { grid-column: 2; }
    .invoice-print-totals > div:not(.invoice-print-amount-words):not(.invoice-print-signature) { display: flex; justify-content: space-between; gap: 12px; }
    .invoice-print-totals > div:not(.invoice-print-amount-words):not(.invoice-print-signature):not(.invoice-print-balance-rule) { padding: 5px 7px; border: 0; border-bottom: 1px solid #dfe6ef; }
    .invoice-print-total-row { font-size: 11px; font-weight: 700; }
    .invoice-print-total-row strong { font-size: 12px; }
    .invoice-print-totals > div:not(.invoice-print-amount-words):not(.invoice-print-signature) > span { color: #667789; font-size: 10px; }
    .invoice-print-amount-words { grid-column: 1; grid-row: 1 / span 8; padding-right: 10px; }
    .invoice-print-amount-line { padding: 6px 0 7px; border-top: 1px solid #d8e0ea; border-bottom: 1px solid #d8e0ea; }
    .invoice-print-amount-line span { display: block; margin-bottom: 3px; color: #5c6b7a; font-size: 9px; letter-spacing: .08em; text-transform: uppercase; }
    .invoice-print-amount-line strong { display: block; color: #17212b; font-size: 10px; font-style: italic; line-height: 1.3; }
    .invoice-print-bank-block { margin-top: 8px; color: #4a5b6d; font-size: 10px; line-height: 1.3; }
    .invoice-print-bank-block p { margin: 0 0 4px; }
    .invoice-print-bank-block div { display: block; margin: 2px 0; }
    .invoice-print-bank-block span { color: #667789; }
    .invoice-print-bank-block strong { color: #17212b; font-weight: 700; }
    .invoice-print-balance { border-top: 2px solid #dfe6ef; border-bottom: 1px solid #dfe6ef; padding-top: 7px; font-size: 13px; }
    .invoice-print-balance-rule { grid-column: 2; height: 0; margin-top: 3px; border-top: 2px solid #17212b; }
    .invoice-print-signature { grid-column: 2; display: block; margin-top: 7px; text-align: center; }
    .invoice-print-signature-space { height: 58px; margin: 4px 0 5px; border-bottom: 1px solid #17212b; display: flex; align-items: end; justify-content: center; overflow: hidden; }
    .invoice-print-signature-image { max-height: 52px; width: auto; max-width: 180px; object-fit: contain; }
    .invoice-print-signature strong { display: block; margin-bottom: 3px; color: #17212b; font-size: 10px; white-space: nowrap; }
    .invoice-print-signature span { color: #17212b; font-size: 10px; font-weight: 700; }
    .invoice-print-footer { margin-top: 16px; border-top: 1px solid #dfe6ef; padding-top: 9px; color: #667789; font-size: 9px; line-height: 1.3; }
    .invoice-print-thank-you { margin-bottom: 4px; color: #17212b; font-weight: 700; }
    .invoice-print-payment-note { font-size: 9px; line-height: 1.3; }
    .invoice-print-page-number { display: none; }
    .invoice-print-page + .invoice-print-page { margin-top: 18px; page-break-before: always; }
    @media print { body { background: #fff; } .invoice-print-page { margin: 0; box-shadow: none; border-radius: 0; } .invoice-print-page + .invoice-print-page { margin-top: 0; page-break-before: always; } }
  `;
}

function openInvoicePrintPreview(invoiceId) {
  const invoice = (state.invoices || []).find((item) => item.id === invoiceId);
  if (!invoice) {
    showToast("Invoice not found.");
    return;
  }
  const modal = document.getElementById("invoicePreviewModal");
  const content = document.getElementById("invoicePreviewContent");
  const copiesSelect = document.getElementById("invoicePrintCopies");
  if (!modal || !content) return;
  if (copiesSelect) {
    copiesSelect.value = String(state.ui.invoicePrintCopies || 1);
  }
  content.innerHTML = getInvoiceDocumentMarkup(invoice, { copyCount: state.ui.invoicePrintCopies || 1 });
  modal.dataset.invoiceId = invoice.id;
  modal.classList.remove("hidden");
  modal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeInvoicePrintPreview() {
  const modal = document.getElementById("invoicePreviewModal");
  const content = document.getElementById("invoicePreviewContent");
  if (!modal || !content) return;
  modal.classList.add("hidden");
  modal.dataset.invoiceId = "";
  modal.setAttribute("aria-hidden", "true");
  content.innerHTML = "";
  document.body.classList.remove("modal-open");
}

function printInvoiceById(invoiceId, copyCount = state.ui.invoicePrintCopies || 1) {
  const invoice = (state.invoices || []).find((item) => item.id === invoiceId);
  if (!invoice) {
    showToast("Invoice not found.");
    return;
  }
  const printWindow = window.open("", "_blank", "width=1080,height=900");
  if (!printWindow) {
    showToast("Popup blocked. Allow popups to print invoices.");
    return;
  }
  printWindow.document.write(`
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>${escapeHtml(invoice.invoiceNumber)}</title>
        <style>${getInvoicePrintStyles()}</style>
      </head>
      <body>${getInvoiceDocumentMarkup(invoice, { copyCount })}</body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.onload = () => {
    printWindow.print();
  };
}

function findInvoiceUsingShow(showId, excludeInvoiceId = "") {
  if (!showId) return null;
  return (state.invoices || []).find((invoice) => invoice.id !== excludeInvoiceId && (invoice.lineItems || []).some((item) => parseInvoiceShowIds(item.showId).includes(showId))) || null;
}

function renderInvoicesPanel() {
  const panel = document.getElementById("invoicesPanel");
  const invoices = sortInvoices(state.invoices || []);
  const clientMasterOptions = getSortedClients();
  const clientOptions = getInvoiceClientOptions(invoices);
  const invoiceYearOptions = ["all", ...getInvoiceYearOptions(invoices)];
  if (!invoiceYearOptions.includes(state.ui.invoiceExportYear)) {
    state.ui.invoiceExportYear = "all";
  }
  const invoiceMonthKeys = getInvoiceMonthOptions(invoices, state.ui.invoiceExportYear || "all");
  const invoiceMonthOptions = [
    { value: "all", label: "All Months" },
    ...invoiceMonthKeys.map((monthKey) => ({
      value: monthKey,
      label: monthGroupLabel(`${monthKey}-01`)
    }))
  ];
  if (state.ui.invoiceExportMonth !== "all" && !invoiceMonthKeys.includes(state.ui.invoiceExportMonth)) {
    state.ui.invoiceExportMonth = "all";
  }
  const lightDesignerOptions = ["all", ...getInvoiceLightDesignerOptions(invoices)];
  if (!lightDesignerOptions.includes(state.ui.invoiceLightDesignerFilter)) {
    state.ui.invoiceLightDesignerFilter = "all";
  }
  const filteredInvoices = filterInvoices(invoices);
  const invoicePagination = getPaginationSlice(filteredInvoices, "invoiceRegisterPage", "invoiceRegisterPageSize");
  const todayKey = dateKey(new Date());
  const nextSevenKey = dateKey(addDays(new Date(), 7));
  const registerInvoices = filteredInvoices;
  const overdueInvoices = registerInvoices
    .filter((invoice) => Number(invoice.balanceDue || 0) > 0 && invoice.dueDate && invoice.dueDate < todayKey && invoice.status !== "cancelled")
    .sort((a, b) => String(a.dueDate || "").localeCompare(String(b.dueDate || "")));
  const dueSoonInvoices = registerInvoices
    .filter((invoice) => Number(invoice.balanceDue || 0) > 0 && invoice.dueDate && invoice.dueDate >= todayKey && invoice.dueDate <= nextSevenKey && invoice.status !== "cancelled")
    .sort((a, b) => String(a.dueDate || "").localeCompare(String(b.dueDate || "")));
  const partialInvoices = registerInvoices
    .filter((invoice) => getInvoicePaymentBucket(invoice) === "partiallyPaid")
    .sort((a, b) => Number(b.balanceDue || 0) - Number(a.balanceDue || 0));
  const recentCollections = registerInvoices
    .flatMap((invoice) => (Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : []).map((payment) => ({
      ...payment,
      invoiceNumber: invoice.invoiceNumber,
      clientName: getInvoiceClientLabel(invoice),
      invoiceId: invoice.id
    })))
    .sort((a, b) => String(b.paymentDate || "").localeCompare(String(a.paymentDate || "")))
    .slice(0, 5);
  const topClientOutstanding = [...registerInvoices.reduce((map, invoice) => {
    const clientLabel = getInvoiceClientLabel(invoice);
    const current = map.get(clientLabel) || { clientLabel, outstanding: 0, invoices: 0 };
    current.outstanding += Number(invoice.balanceDue || 0);
    current.invoices += 1;
    map.set(clientLabel, current);
    return map;
  }, new Map()).values()]
    .filter((item) => item.outstanding > 0)
    .sort((a, b) => b.outstanding - a.outstanding)
    .slice(0, 5);
  const availableShows = [...state.shows].sort((a, b) => getShowStartDate(b).localeCompare(getShowStartDate(a)));
  const editingInvoice = state.ui.editingInvoiceId
    ? invoices.find((invoice) => invoice.id === state.ui.editingInvoiceId)
    : null;
  const activeInvoiceSubtab = editingInvoice ? "create" : state.ui.invoiceSubtab || "create";
  const paymentRows = getInvoiceReconciliationRows(invoices);
  const paymentReconYearOptions = getPaymentReconciliationYearOptions(paymentRows);
  if (!["all", ...paymentReconYearOptions].includes(state.ui.paymentReconYear)) {
    state.ui.paymentReconYear = "all";
  }
  const paymentReconMonthKeys = getPaymentReconciliationMonthOptions(paymentRows, state.ui.paymentReconYear || "all");
  if (state.ui.paymentReconMonth !== "all" && !paymentReconMonthKeys.includes(state.ui.paymentReconMonth)) {
    state.ui.paymentReconMonth = "all";
  }
  const paymentReconClientOptions = ["all", ...new Set(paymentRows.map((row) => row.clientLabel).filter(Boolean))].sort((a, b) => {
    if (a === "all") return -1;
    if (b === "all") return 1;
    return a.localeCompare(b);
  });
  if (!paymentReconClientOptions.includes(state.ui.paymentReconClient)) {
    state.ui.paymentReconClient = "all";
  }
  const filteredPaymentRows = filterInvoiceReconciliationRows(paymentRows);
  const paymentReconPagination = getPaginationSlice(filteredPaymentRows, "paymentReconPage", "paymentReconPageSize");
  const paymentReconSummary = {
    totalCollected: filteredPaymentRows.reduce((sum, row) => sum + Number(row.amount || 0), 0),
    totalPayments: filteredPaymentRows.length,
    clientsCovered: new Set(filteredPaymentRows.map((row) => row.clientLabel).filter(Boolean)).size,
    invoicesCovered: new Set(filteredPaymentRows.map((row) => row.invoiceId).filter(Boolean)).size
  };
  const draftShows = (state.ui.invoiceDraftShowIds || [])
    .map((showId) => state.shows.find((show) => show.id === showId))
    .filter(Boolean);
  const summary = invoices.reduce((acc, invoice) => {
    const label = getInvoiceStatusLabel(invoice);
    acc.total += 1;
    acc.balanceDue += Number(invoice.balanceDue || 0);
    if (label === "Paid") acc.paid += 1;
    if (label === "Overdue") acc.overdue += 1;
    if (label === "Draft") acc.draft += 1;
    return acc;
  }, { total: 0, paid: 0, overdue: 0, draft: 0, balanceDue: 0 });

  panel.innerHTML = `
    <div class="stack">
      <div class="form-header">
        <div>
          <h3>Invoices</h3>
          <p class="muted-note">Create invoice records, bundle multiple shows, and track payment status in one place.</p>
        </div>
        <div class="toolbar">
          <span class="pill">${summary.total} ${summary.total === 1 ? "invoice" : "invoices"}</span>
        </div>
      </div>
      <div class="summary-grid">
        <div class="summary-card">
          <span class="summary-kicker">Outstanding</span>
          <strong>${formatCurrency(summary.balanceDue)}</strong>
          <span class="summary-foot">Current unpaid balance across all invoices.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Paid</span>
          <strong>${summary.paid}</strong>
          <span class="summary-foot">Invoices fully settled.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Overdue</span>
          <strong>${summary.overdue}</strong>
          <span class="summary-foot">Invoices past due with balance still open.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Drafts</span>
          <strong>${summary.draft}</strong>
          <span class="summary-foot">Invoices not yet sent.</span>
        </div>
      </div>
      <div class="invoice-subtabs" role="tablist" aria-label="Invoice sections">
        <button type="button" class="${activeInvoiceSubtab === "create" ? "is-active" : ""}" data-invoice-subtab="create">${getDirtyTabLabel("Create Invoice", "invoice")}</button>
        <button type="button" class="${activeInvoiceSubtab === "register" ? "is-active" : ""}" data-invoice-subtab="register">Invoice Register</button>
        <button type="button" class="${activeInvoiceSubtab === "payments" ? "is-active" : ""}" data-invoice-subtab="payments">Payments</button>
      </div>
      <div class="stack ${activeInvoiceSubtab === "create" ? "" : "hidden"}" data-invoice-section="create">
        <div class="form-header">
          <div>
            <h4>${editingInvoice ? `Editing ${editingInvoice.invoiceNumber}` : "Create Invoice"}</h4>
            <p class="muted-note">Link shows where useful, or add manual billing lines for travel, rental, or other work.</p>
          </div>
          <div class="toolbar">
            ${editingInvoice ? '<span class="pill edit-pill">Edit Mode</span>' : ""}
            <button type="button" class="secondary small" id="newInvoiceButton">New Invoice</button>
          </div>
        </div>
        <form id="invoiceForm" class="stack tight editor-form" autocomplete="off">
          <input type="hidden" name="invoiceId">
          <div class="form-grid editor-section">
            <label class="field"><span>Invoice Number</span><input type="text" name="invoiceNumber" required autocomplete="off" data-form-type="other"></label>
            <label class="field">
              <span>Client</span>
              <select name="clientId" required data-searchable="true" data-search-placeholder="Search clients">
                <option value="">Select client</option>
                ${clientMasterOptions.map((client) => `<option value="${client.id}">${escapeHtml(getClientDisplayName(client))}</option>`).join("")}
              </select>
            </label>
            <label class="field"><span>Issue Date</span><input type="date" name="issueDate" required autocomplete="off"></label>
            <label class="field">
              <span>Status</span>
              <select name="status">
                <option value="draft">Draft</option>
                <option value="sent">Sent</option>
                <option value="partially_paid">Partially Paid</option>
                <option value="paid">Paid</option>
                <option value="cancelled">Cancelled</option>
              </select>
            </label>
            <label class="field">
              <span>Payment Terms</span>
              <select name="paymentTerms">
                <option value="Due on receipt">Due on receipt</option>
                <option value="Net 10">Net 10</option>
                <option value="Net 15">Net 15</option>
                <option value="Net 30">Net 30</option>
              </select>
            </label>
            <label class="field"><span>Amount Paid</span><input type="number" name="amountPaid" min="0" step="0.01" autocomplete="off"></label>
          </div>
          <label class="field editor-section"><span>Notes</span><input type="text" name="notes" autocomplete="off"></label>
          <div class="stack tight editor-section">
            <div class="form-header">
              <div>
                <h4>Particulars</h4>
              </div>
              <button type="button" class="secondary small" id="addInvoiceLineItem">Add Line Item</button>
            </div>
            <div id="invoiceLineItems" class="stack tight"></div>
          </div>
          <div class="assignment-card" id="invoiceTotalsCard">
            <div class="meta">Subtotal: <strong id="invoiceSubtotalPreview">${formatCurrency(0)}</strong></div>
            <div class="meta">Discount: <strong id="invoiceDiscountPreview">-</strong></div>
            <div class="meta">SGST (9%): <strong id="invoiceSgstPreview">-</strong></div>
            <div class="meta">CGST (9%): <strong id="invoiceCgstPreview">-</strong></div>
            <div class="meta">IGST (18%): <strong id="invoiceIgstPreview">-</strong></div>
            <div class="meta">Total: <strong id="invoiceTotalPreview">${formatCurrency(0)}</strong></div>
            <div class="meta">Balance Due: <strong id="invoiceBalancePreview">${formatCurrency(0)}</strong></div>
          </div>
          <div class="toolbar editor-actions">
            <button type="submit">${editingInvoice ? "Update Invoice" : "Save Invoice"}</button>
            <button type="button" class="ghost" id="cancelInvoiceEdit">${editingInvoice ? "Cancel Edit" : "Clear"}</button>
            ${editingInvoice ? '<button type="button" class="danger" id="deleteInvoiceButton">Delete Invoice</button>' : ""}
          </div>
          <div id="invoiceFormMessage" class="message"></div>
        </form>
      </div>
      <div class="stack ${activeInvoiceSubtab === "register" ? "" : "hidden"}" data-invoice-section="register">
        <div class="form-header">
          <div>
            <h4>Invoice Register</h4>
            <p class="muted-note">Recent invoices with status, balance, and linked shows.</p>
          </div>
          <span class="pill">${filteredInvoices.length} ${filteredInvoices.length === 1 ? "result" : "results"}</span>
        </div>
        <div class="client-detail-grid invoice-collections-grid">
          <section class="assignment-card">
            <strong>Overdue</strong>
            <div class="stack tight client-detail-list">
              ${overdueInvoices.length ? overdueInvoices.slice(0, 5).map((invoice) => `
                <div class="client-detail-row">
                  <div>
                    <strong>${escapeHtml(invoice.invoiceNumber)}</strong>
                    <div class="meta">${escapeHtml(getInvoiceClientLabel(invoice))} · Due ${escapeHtml(formatInvoiceDate(invoice.dueDate))}</div>
                  </div>
                  <strong>${formatCurrency(invoice.balanceDue || 0)}</strong>
                </div>
              `).join("") : '<p class="meta">No overdue invoices in this view.</p>'}
            </div>
          </section>
          <section class="assignment-card">
            <strong>Due In 7 Days</strong>
            <div class="stack tight client-detail-list">
              ${dueSoonInvoices.length ? dueSoonInvoices.slice(0, 5).map((invoice) => `
                <div class="client-detail-row">
                  <div>
                    <strong>${escapeHtml(invoice.invoiceNumber)}</strong>
                    <div class="meta">${escapeHtml(getInvoiceClientLabel(invoice))} · Due ${escapeHtml(formatInvoiceDate(invoice.dueDate))}</div>
                  </div>
                  <strong>${formatCurrency(invoice.balanceDue || 0)}</strong>
                </div>
              `).join("") : '<p class="meta">Nothing due in the next 7 days.</p>'}
            </div>
          </section>
          <section class="assignment-card">
            <strong>Partially Paid</strong>
            <div class="stack tight client-detail-list">
              ${partialInvoices.length ? partialInvoices.slice(0, 5).map((invoice) => `
                <div class="client-detail-row">
                  <div>
                    <strong>${escapeHtml(invoice.invoiceNumber)}</strong>
                    <div class="meta">${escapeHtml(getInvoiceClientLabel(invoice))} · Paid ${formatCurrency(invoice.amountPaid || 0)}</div>
                  </div>
                  <strong>${formatCurrency(invoice.balanceDue || 0)}</strong>
                </div>
              `).join("") : '<p class="meta">No partially paid invoices in this view.</p>'}
            </div>
          </section>
          <section class="assignment-card">
            <strong>Recent Collections</strong>
            <div class="stack tight client-detail-list">
              ${recentCollections.length ? recentCollections.map((payment) => `
                <div class="client-detail-row">
                  <div>
                    <strong>${escapeHtml(formatInvoiceDate(payment.paymentDate))}</strong>
                    <div class="meta">${escapeHtml(payment.clientName)} · ${escapeHtml(payment.invoiceNumber)}</div>
                  </div>
                  <strong>${formatCurrency(payment.amount || 0)}</strong>
                </div>
              `).join("") : '<p class="meta">No recent collections in this view.</p>'}
            </div>
          </section>
          <section class="assignment-card">
            <strong>Top Outstanding Clients</strong>
            <div class="stack tight client-detail-list">
              ${topClientOutstanding.length ? topClientOutstanding.map((entry) => `
                <div class="client-detail-row">
                  <div>
                    <strong>${escapeHtml(entry.clientLabel)}</strong>
                    <div class="meta">${entry.invoices} ${entry.invoices === 1 ? "invoice" : "invoices"} in this view</div>
                  </div>
                  <strong>${formatCurrency(entry.outstanding)}</strong>
                </div>
              `).join("") : '<p class="meta">No outstanding client balances in this view.</p>'}
            </div>
          </section>
        </div>
        <div class="shows-toolbar invoice-toolbar">
          <div class="shows-toolbar-top">
            <label class="sort-control invoice-search-control">
              <span>Search</span>
              <input type="search" id="invoiceSearchInput" placeholder="Invoice, client, show, line item" value="${escapeHtml(state.ui.invoiceSearchQuery || "")}" autocomplete="off">
            </label>
            <button type="button" class="secondary search-submit-button" id="applyInvoiceSearchButton">Search</button>
          </div>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Status</span>
              <select id="invoiceStatusFilter">
                <option value="all" ${state.ui.invoiceStatusFilter === "all" ? "selected" : ""}>All Statuses</option>
                <option value="draft" ${state.ui.invoiceStatusFilter === "draft" ? "selected" : ""}>Draft</option>
                <option value="sent" ${state.ui.invoiceStatusFilter === "sent" ? "selected" : ""}>Sent</option>
                <option value="overdue" ${state.ui.invoiceStatusFilter === "overdue" ? "selected" : ""}>Overdue</option>
                <option value="cancelled" ${state.ui.invoiceStatusFilter === "cancelled" ? "selected" : ""}>Cancelled</option>
              </select>
            </label>
            <label class="sort-control">
              <span>Payment</span>
              <select id="invoicePaymentFilter">
                <option value="all" ${state.ui.invoicePaymentFilter === "all" ? "selected" : ""}>All Payments</option>
                <option value="unpaid" ${state.ui.invoicePaymentFilter === "unpaid" ? "selected" : ""}>Unpaid</option>
                <option value="paid" ${state.ui.invoicePaymentFilter === "paid" ? "selected" : ""}>Paid</option>
                <option value="partiallyPaid" ${state.ui.invoicePaymentFilter === "partiallyPaid" ? "selected" : ""}>Partially Paid</option>
              </select>
            </label>
            <label class="sort-control">
              <span>Client</span>
              <select id="invoiceClientFilter">
                <option value="all" ${state.ui.invoiceClientFilter === "all" ? "selected" : ""}>All Clients</option>
                ${clientOptions.map((client) => `<option value="${client}" ${state.ui.invoiceClientFilter === client ? "selected" : ""}>${client}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="shows-toolbar-top invoice-toolbar-secondary">
            <label class="sort-control">
              <span>Year</span>
              <select id="invoiceYearFilter">
                ${invoiceYearOptions.map((year) => `<option value="${year}" ${state.ui.invoiceExportYear === year ? "selected" : ""}>${year === "all" ? "All Years" : year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="invoiceMonthFilter">
                ${invoiceMonthOptions.map((option) => `<option value="${option.value}" ${state.ui.invoiceExportMonth === option.value ? "selected" : ""}>${option.label}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Light Designer</span>
              <select id="invoiceLightDesignerFilter">
                ${lightDesignerOptions.map((name) => `<option value="${name}" ${state.ui.invoiceLightDesignerFilter === name ? "selected" : ""}>${name === "all" ? "All Light Designers" : name}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Sort By</span>
              <select id="invoiceSortMode">
                <option value="issueDate" ${state.ui.invoiceSortMode === "issueDate" ? "selected" : ""}>Issue Date</option>
                <option value="dueDate" ${state.ui.invoiceSortMode === "dueDate" ? "selected" : ""}>Due Date</option>
                <option value="client" ${state.ui.invoiceSortMode === "client" ? "selected" : ""}>Client</option>
                <option value="lightDesigner" ${state.ui.invoiceSortMode === "lightDesigner" ? "selected" : ""}>Light Designer</option>
              </select>
            </label>
          </div>
        </div>
        <div class="approval-list">
          ${invoicePagination.items.length ? invoicePagination.items.map((invoice) => `
            <article class="show-card">
              <header>
                <div>
                  <h4>${invoice.invoiceNumber}</h4>
                  <div class="meta">${escapeHtml(getClientDisplayValue(invoice.clientId, invoice.clientName))} · Issued ${formatInvoiceDate(invoice.issueDate)}${invoice.dueDate ? ` · Due ${formatInvoiceDate(invoice.dueDate)}` : ""}</div>
                  ${getInvoiceLinkedShowNames(invoice).length ? `<div class="meta">Show: ${escapeHtml(getInvoiceLinkedShowNames(invoice).join(", "))}</div>` : ""}
                  <div class="meta">${getInvoiceLightDesignerLabel(invoice) ? `Light Designer: ${getInvoiceLightDesignerLabel(invoice)}` : "Light Designer: -"}</div>
                </div>
                <div class="toolbar">
                  <span class="pill invoice-status-pill invoice-status-${getInvoiceStatusTone(invoice)}">${getInvoiceStatusLabel(invoice)}</span>
                  <button type="button" class="secondary small" data-mark-payment="${invoice.id}">Mark Payment</button>
                  <button type="button" class="ghost small" data-duplicate-invoice="${invoice.id}">Duplicate</button>
                  <button type="button" class="ghost small" data-preview-invoice="${invoice.id}">Preview</button>
                  <button type="button" class="secondary small" data-edit-invoice="${invoice.id}">Edit</button>
                </div>
              </header>
              <div class="show-banner">
                <span class="show-banner-item">${formatCurrency(invoice.totalAmount)}</span>
                <span class="show-banner-item">Paid ${formatCurrency(invoice.amountPaid)}</span>
                <span class="show-banner-item">Balance ${formatCurrency(invoice.balanceDue)}</span>
                <span class="show-banner-item">${invoice.lineItems.length} ${invoice.lineItems.length === 1 ? "line" : "lines"}</span>
              </div>
              ${state.ui.markPaymentInvoiceId === invoice.id ? `
                <form class="invoice-payment-form" data-payment-form="${invoice.id}">
                  <label class="field"><span>Payment Date</span><input type="date" name="paymentDate" value="${dateKey(new Date())}" required></label>
                  <label class="field"><span>Amount</span><input type="number" name="amount" min="0.01" step="0.01" max="${Number(invoice.balanceDue || 0)}" required></label>
                  <label class="field invoice-payment-note"><span>Note</span><input type="text" name="note" placeholder="Reference, UTR, cash, etc." autocomplete="off"></label>
                  <div class="toolbar">
                    <button type="submit" class="secondary small">Save Payment</button>
                    <button type="button" class="ghost small" data-cancel-payment="${invoice.id}">Cancel</button>
                  </div>
                  <div class="message" data-payment-message></div>
                </form>
              ` : ""}
              ${getInvoicePaymentHistoryMarkup(invoice)}
            </article>
          `).join("") : invoices.length ? "<p>No invoices match the current filters.</p>" : "<p>No invoices yet. Create the first invoice above.</p>"}
        </div>
        ${renderPaginationControls("invoice-register", invoicePagination, "invoices")}
      </div>
      <div class="stack ${activeInvoiceSubtab === "payments" ? "" : "hidden"}" data-invoice-section="payments">
        <div class="form-header">
          <div>
            <h4>Payment Reconciliation</h4>
            <p class="muted-note">Track all received payments across invoices in one place.</p>
          </div>
          <span class="pill">${filteredPaymentRows.length} ${filteredPaymentRows.length === 1 ? "payment" : "payments"}</span>
        </div>
        <div class="summary-grid">
          <div class="summary-card">
            <span class="summary-kicker">Collected</span>
            <strong>${formatCurrency(paymentReconSummary.totalCollected)}</strong>
            <span class="summary-foot">Total receipts in this view.</span>
          </div>
          <div class="summary-card">
            <span class="summary-kicker">Payments</span>
            <strong>${paymentReconSummary.totalPayments}</strong>
            <span class="summary-foot">Recorded payment entries.</span>
          </div>
          <div class="summary-card">
            <span class="summary-kicker">Clients</span>
            <strong>${paymentReconSummary.clientsCovered}</strong>
            <span class="summary-foot">Clients covered by these receipts.</span>
          </div>
          <div class="summary-card">
            <span class="summary-kicker">Invoices</span>
            <strong>${paymentReconSummary.invoicesCovered}</strong>
            <span class="summary-foot">Invoices touched by these receipts.</span>
          </div>
        </div>
        <div class="shows-toolbar invoice-toolbar">
          <div class="shows-toolbar-top">
            <label class="sort-control invoice-search-control">
              <span>Search</span>
              <input type="search" id="paymentReconSearchInput" placeholder="Invoice, client, note, show" value="${escapeHtml(state.ui.paymentReconSearchQuery || "")}" autocomplete="off">
            </label>
            <button type="button" class="secondary search-submit-button" id="applyPaymentReconSearchButton">Search</button>
          </div>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Client</span>
              <select id="paymentReconClientFilter">
                ${paymentReconClientOptions.map((option) => `<option value="${option}" ${state.ui.paymentReconClient === option ? "selected" : ""}>${option === "all" ? "All Clients" : escapeHtml(option)}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Year</span>
              <select id="paymentReconYearFilter">
                <option value="all" ${state.ui.paymentReconYear === "all" ? "selected" : ""}>All Years</option>
                ${paymentReconYearOptions.map((year) => `<option value="${year}" ${state.ui.paymentReconYear === year ? "selected" : ""}>${year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="paymentReconMonthFilter">
                <option value="all" ${state.ui.paymentReconMonth === "all" ? "selected" : ""}>All Months</option>
                ${paymentReconMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.paymentReconMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
              </select>
            </label>
          </div>
        </div>
        <div class="client-ledger-table-wrap">
          <table class="client-ledger-table">
            <thead>
              <tr>
                <th>Payment Date</th>
                <th>Invoice</th>
                <th>Client</th>
                <th>Amount</th>
                <th>Note</th>
                <th>Status</th>
                <th>Balance</th>
                <th>Light Designer</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              ${paymentReconPagination.items.length ? paymentReconPagination.items.map((row) => `
                <tr>
                  <td>${escapeHtml(formatInvoiceDate(row.paymentDate))}</td>
                  <td>
                    <strong>${escapeHtml(row.invoiceNumber)}</strong>
                    ${row.showNames ? `<div class="meta">${escapeHtml(row.showNames)}</div>` : ""}
                  </td>
                  <td>${escapeHtml(row.clientLabel)}</td>
                  <td><strong>${formatCurrency(row.amount)}</strong></td>
                  <td>${escapeHtml(row.note || "-")}</td>
                  <td><span class="pill invoice-status-pill invoice-status-${getInvoiceStatusTone({ status: row.invoiceRawStatus, amountPaid: row.amountPaid, balanceDue: row.balanceDue })}">${escapeHtml(row.invoiceStatus)}</span></td>
                  <td>${formatCurrency(row.balanceDue)}</td>
                  <td>${escapeHtml(row.lightDesigner || "-")}</td>
                  <td><button type="button" class="secondary small" data-open-payment-invoice="${row.invoiceId}">Open Invoice</button></td>
                </tr>
              `).join("") : `<tr><td colspan="9">No payments match the current filters.</td></tr>`}
            </tbody>
          </table>
        </div>
        ${renderPaginationControls("invoice-payments", paymentReconPagination, "payments")}
      </div>
    </div>
  `;

  enhanceCustomSelects(panel);

  const form = document.getElementById("invoiceForm");
  const lineItemsContainer = document.getElementById("invoiceLineItems");
  const clientSelect = form.elements.namedItem("clientId");

  function getInvoiceShowPickerLabel(selectedShowIds = []) {
    if (!selectedShowIds.length) return "No linked show - custom details";
    const names = selectedShowIds
      .map((showId) => state.shows.find((show) => show.id === showId)?.showName)
      .filter(Boolean);
    if (!names.length) return "No linked show - custom details";
    return names.length === 1 ? names[0] : `${names.length} shows selected`;
  }

  function getInvoiceShowPickerMarkup(selectedClientId = "", selectedShowIds = []) {
    const uniqueSelectedShowIds = parseInvoiceShowIds(selectedShowIds);
    if (!selectedClientId) {
      return `
        <div class="invoice-show-picker is-disabled" data-show-picker>
          <input type="hidden" name="lineShowId" value="">
          <button type="button" class="invoice-show-picker-trigger" disabled>Select client to link shows</button>
        </div>
      `;
    }
    const filteredShows = availableShows.filter((show) => show.clientId === selectedClientId);
    const selectedOutsideClientShows = uniqueSelectedShowIds
      .map((showId) => state.shows.find((show) => show.id === showId))
      .filter((show) => show && show.clientId !== selectedClientId);
    const pickerShows = [...selectedOutsideClientShows, ...filteredShows.filter((show) => !selectedOutsideClientShows.some((selectedShow) => selectedShow.id === show.id))];
    return `
      <div class="invoice-show-picker" data-show-picker>
        <input type="hidden" name="lineShowId" value="${serializeInvoiceShowIds(uniqueSelectedShowIds)}">
        <button type="button" class="invoice-show-picker-trigger" data-show-picker-trigger>${escapeHtml(getInvoiceShowPickerLabel(uniqueSelectedShowIds))}</button>
        <div class="invoice-show-picker-menu hidden" data-show-picker-menu>
          <input type="search" class="invoice-show-picker-search" data-show-picker-search placeholder="Search shows" autocomplete="off">
          <label class="invoice-show-picker-option">
            <input type="checkbox" value="" data-clear-show-selection ${uniqueSelectedShowIds.length ? "" : "checked"}>
            <span>No linked show - custom details</span>
          </label>
          ${pickerShows.map((show) => {
      const duplicateInvoice = findInvoiceUsingShow(show.id, editingInvoice?.id || "");
      const warning = duplicateInvoice ? " [Already invoiced]" : "";
            const outsideClient = show.clientId !== selectedClientId ? " [Different client]" : "";
            return `
              <label class="invoice-show-picker-option" data-show-picker-searchable>
                <input type="checkbox" value="${show.id}" ${uniqueSelectedShowIds.includes(show.id) ? "checked" : ""} data-show-picker-checkbox>
                <span>${show.showName} · ${formatDateRange(getShowStartDate(show), getShowEndDate(show))}${warning}${outsideClient}</span>
              </label>
            `;
          }).join("") || '<div class="meta">No shows for selected client.</div>'}
        </div>
      </div>
    `;
  }

  function refreshLineItemShowOptions() {
    const selectedClientId = clientSelect.value;
    [...lineItemsContainer.querySelectorAll("[data-show-picker]")].forEach((picker) => {
      const selectedShowIds = parseInvoiceShowIds(picker.querySelector('input[name="lineShowId"]')?.value || "");
      picker.outerHTML = getInvoiceShowPickerMarkup(selectedClientId, selectedShowIds);
    });
    wireInvoiceShowPickers();
  }

  function syncInvoiceShowPicker(picker) {
    const hiddenInput = picker.querySelector('input[name="lineShowId"]');
    const selectedShowIds = [...picker.querySelectorAll("[data-show-picker-checkbox]:checked")].map((checkbox) => checkbox.value).filter(Boolean);
    hiddenInput.value = serializeInvoiceShowIds(selectedShowIds);
    const clearOption = picker.querySelector("[data-clear-show-selection]");
    if (clearOption) {
      clearOption.checked = !selectedShowIds.length;
    }
    const trigger = picker.querySelector("[data-show-picker-trigger]");
    if (trigger) {
      trigger.textContent = getInvoiceShowPickerLabel(selectedShowIds);
    }
  }

  function wireInvoiceShowPickers() {
    lineItemsContainer.querySelectorAll("[data-show-picker]").forEach((picker) => {
      if (picker.dataset.wired === "true") return;
      picker.dataset.wired = "true";
      const trigger = picker.querySelector("[data-show-picker-trigger]");
      const menu = picker.querySelector("[data-show-picker-menu]");
      trigger?.addEventListener("click", () => {
        menu?.classList.toggle("hidden");
        if (!menu?.classList.contains("hidden")) {
          menu.querySelector("[data-show-picker-search]")?.focus();
        }
      });
      picker.querySelector("[data-show-picker-search]")?.addEventListener("input", (event) => {
        const query = event.currentTarget.value.trim().toLowerCase();
        picker.querySelectorAll("[data-show-picker-searchable]").forEach((option) => {
          option.classList.toggle("hidden", Boolean(query) && !option.textContent.toLowerCase().includes(query));
        });
      });
      picker.querySelector("[data-clear-show-selection]")?.addEventListener("change", (event) => {
        if (!event.currentTarget.checked) return;
        picker.querySelectorAll("[data-show-picker-checkbox]").forEach((checkbox) => {
          checkbox.checked = false;
        });
        setDirtyForm("invoice");
        syncInvoiceShowPicker(picker);
        const row = picker.closest("[data-invoice-line-item]");
        row?.querySelector('input[name="lineDescription"]')?.focus();
        updateTotalsPreview();
      });
      picker.querySelectorAll("[data-show-picker-checkbox]").forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const selectedShow = state.shows.find((show) => show.id === checkbox.value);
          const duplicateInvoice = checkbox.checked ? findInvoiceUsingShow(checkbox.value, editingInvoice?.id || "") : null;
          if (selectedShow && duplicateInvoice) {
            const proceed = window.confirm(
              `"${selectedShow.showName}" is already linked to invoice ${duplicateInvoice.invoiceNumber}. Do you still want to add it here?`
            );
            if (!proceed) {
              checkbox.checked = false;
              syncInvoiceShowPicker(picker);
              updateTotalsPreview();
              return;
            }
          }
          syncInvoiceShowPicker(picker);
          const row = picker.closest("[data-invoice-line-item]");
          const selectedShowIds = parseInvoiceShowIds(picker.querySelector('input[name="lineShowId"]')?.value || "");
          const selectedShows = selectedShowIds.map((showId) => state.shows.find((show) => show.id === showId)).filter(Boolean);
          const quantityInput = row?.querySelector('input[name="lineQuantity"]');
          const descriptionInput = row?.querySelector('input[name="lineDescription"]');
          if (selectedShows.length && quantityInput) {
            quantityInput.value = selectedShows.length;
          }
          if (selectedShows.length && descriptionInput && !descriptionInput.value.trim()) {
            descriptionInput.value = selectedShows.map((show) => show.showName).join(", ");
          }
          const duplicateWarning = row?.querySelector("[data-duplicate-warning]");
          const duplicates = selectedShows
            .map((show) => findInvoiceUsingShow(show.id, editingInvoice?.id || ""))
            .filter(Boolean);
          if (duplicateWarning) {
            duplicateWarning.classList.toggle("hidden", !duplicates.length);
            duplicateWarning.querySelector("strong").textContent = duplicates.length ? `Already on ${[...new Set(duplicates.map((invoice) => invoice.invoiceNumber))].join(", ")}` : "";
          }
          setDirtyForm("invoice");
          updateTotalsPreview();
        });
      });
      syncInvoiceShowPicker(picker);
    });
  }

  function updateTotalsPreview() {
    const amountPaid = Number(form.elements.namedItem("amountPaid").value || 0);
    const rows = [...lineItemsContainer.querySelectorAll("[data-invoice-line-item]")];
    const lineItems = rows.map((row) => {
      const quantity = Number(row.querySelector('input[name="lineQuantity"]').value || 0);
      const unitRate = Number(row.querySelector('input[name="lineUnitRate"]').value || 0);
      const discount = row.querySelector('input[name="lineDiscount"]').value || "";
      return { quantity: Math.max(quantity || 1, 1), unitRate, discount };
    });
    const subtotal = lineItems.reduce((sum, item, index) => {
      const amount = item.quantity * item.unitRate;
      const discountAmount = getDiscountAmount(item.discount, amount);
      const netAmount = Math.max(0, amount - discountAmount);
      const row = rows[index];
      const amountNode = row.querySelector("[data-line-amount]");
      if (amountNode) {
        amountNode.textContent = formatCurrency(netAmount);
      }
      return sum + amount;
    }, 0);
    const selectedClient = getClientById(clientSelect.value);
    const calculation = getInvoiceCalculationFromValues(lineItems, selectedClient?.state || "");
    const totalAmount = calculation.totalAmount;
    const balanceDue = Math.max(0, totalAmount - amountPaid);
    document.getElementById("invoiceSubtotalPreview").textContent = formatCurrency(subtotal);
    document.getElementById("invoiceDiscountPreview").textContent = calculation.discountAmount ? formatCurrency(calculation.discountAmount) : "-";
    document.getElementById("invoiceSgstPreview").textContent = calculation.sgstAmount ? formatCurrency(calculation.sgstAmount) : "-";
    document.getElementById("invoiceCgstPreview").textContent = calculation.cgstAmount ? formatCurrency(calculation.cgstAmount) : "-";
    document.getElementById("invoiceIgstPreview").textContent = calculation.igstAmount ? formatCurrency(calculation.igstAmount) : "-";
    document.getElementById("invoiceTotalPreview").textContent = formatCurrency(totalAmount);
    document.getElementById("invoiceBalancePreview").textContent = formatCurrency(balanceDue);
  }

  function addLineItemRow(lineItem = null) {
    const row = document.createElement("div");
    row.className = "form-grid invoice-line-item-row";
    row.dataset.invoiceLineItem = "true";
    const lineId = lineItem?.id || uid("line");
    const selectedShowIds = parseInvoiceShowIds(lineItem?.showId || "");
    const duplicateInvoices = selectedShowIds.map((showId) => findInvoiceUsingShow(showId, editingInvoice?.id || "")).filter(Boolean);
    const selectedClientId = clientSelect.value || lineItem?.clientId || "";
    const rememberedDescription = state.ui.invoiceLineDefaults?.description || "";
    const rememberedSac = state.ui.invoiceLineDefaults?.sac || "";
    const lineDescription = lineItem?.description ?? rememberedDescription;
    const lineSac = lineItem?.sac ?? rememberedSac;
    row.innerHTML = `
      <input type="hidden" name="lineId" value="${lineId}">
      <label class="field">
        <span>Particulars</span>
        <input type="text" name="lineDescription" value="${lineDescription}" placeholder="Light Designing and Operating Charges" required autocomplete="off" data-form-type="other">
      </label>
      <label class="field">
        <span>SAC</span>
        <input type="text" name="lineSac" value="${lineSac}" autocomplete="off" data-form-type="other">
      </label>
      <label class="field">
        <span>Qty</span>
        <input type="number" name="lineQuantity" min="0.01" step="0.01" value="${lineItem?.quantity ?? 1}" autocomplete="off">
      </label>
      <label class="field">
        <span>Unit Rate</span>
        <input type="number" name="lineUnitRate" min="0" step="0.01" value="${lineItem?.unitRate ?? 0}" autocomplete="off">
      </label>
      <label class="field invoice-line-wide-field">
        <span>Description</span>
        <input type="text" name="lineCustomDetails" value="${lineItem?.customDetails || ""}" placeholder="Write extra custom details here" autocomplete="off" data-form-type="other">
      </label>
      <label class="field invoice-line-wide-field">
        <span>Select Show(s)</span>
        ${getInvoiceShowPickerMarkup(selectedClientId, selectedShowIds)}
      </label>
      <label class="field invoice-line-wide-field">
        <span>Discount</span>
        <input type="text" name="lineDiscount" value="${lineItem?.discount || ""}" placeholder="Enter amount or %, for example 1000 or 10%" autocomplete="off" data-form-type="other">
      </label>
      <div class="assignment-card">
        <div class="meta">Amount</div>
        <strong data-line-amount>${formatCurrency(lineItem?.amount || 0)}</strong>
      </div>
      <div class="assignment-card invoice-warning-card ${duplicateInvoices.length ? "" : "hidden"}" data-duplicate-warning>
        <div class="meta">Warning</div>
        <strong>${duplicateInvoices.length ? `Already on ${[...new Set(duplicateInvoices.map((invoice) => invoice.invoiceNumber))].join(", ")}` : ""}</strong>
      </div>
      <button type="button" class="ghost small remove-line-item">Remove</button>
    `;
    lineItemsContainer.append(row);
    wireInvoiceShowPickers();
    const descriptionInput = row.querySelector('input[name="lineDescription"]');
    const sacInput = row.querySelector('input[name="lineSac"]');
    const customDetailsInput = row.querySelector('input[name="lineCustomDetails"]');
    const discountInput = row.querySelector('input[name="lineDiscount"]');
    const quantityInput = row.querySelector('input[name="lineQuantity"]');
    const unitRateInput = row.querySelector('input[name="lineUnitRate"]');
    [descriptionInput, sacInput, customDetailsInput, discountInput, quantityInput, unitRateInput].forEach((input) => {
      input.addEventListener("input", updateTotalsPreview);
    });
    row.querySelector(".remove-line-item").addEventListener("click", () => {
      const confirmed = window.confirm("Remove this invoice line item?");
      if (!confirmed) return;
      row.remove();
      setDirtyForm("invoice");
      updateTotalsPreview();
    });
    updateTotalsPreview();
  }

  function quickAddShowAsLineItem(showId) {
    const show = state.shows.find((item) => item.id === showId);
    if (!show) return;
    const duplicateInvoice = findInvoiceUsingShow(show.id, editingInvoice?.id || "");
    if (duplicateInvoice) {
      const proceed = window.confirm(
        `"${show.showName}" is already linked to invoice ${duplicateInvoice.invoiceNumber}. Do you still want to add it here?`
      );
    if (!proceed) return;
    }
    addLineItemRow({
      showId: show.id,
      description: defaultInvoiceLineDescription(show),
      sac: state.ui.invoiceLineDefaults?.sac || "",
      discount: "",
      quantity: 1,
      unitRate: Number(show.amountShow || 0),
      amount: Number(show.amountShow || 0)
    });
    if (!form.elements.namedItem("clientId").value && show.clientId) {
      form.elements.namedItem("clientId").value = show.clientId;
      syncCustomSelect(form.elements.namedItem("clientId"));
    }
    setDirtyForm("invoice");
    updateTotalsPreview();
  }

  document.getElementById("newInvoiceButton")?.addEventListener("click", () => {
    if (!confirmDiscardDirtyForm("start a new invoice")) return;
    clearDirtyForm("invoice");
    resetInvoiceEditingState();
    state.ui.invoiceSubtab = "create";
    saveState(state);
    renderDashboard();
  });

  document.getElementById("cancelInvoiceEdit")?.addEventListener("click", () => {
    if (!confirmDiscardDirtyForm("leave this invoice form")) return;
    clearDirtyForm("invoice");
    resetInvoiceEditingState();
    state.ui.invoiceSubtab = "create";
    saveState(state);
    renderDashboard();
  });

  const invoiceSearchInput = document.getElementById("invoiceSearchInput");
  const applyInvoiceSearch = () => {
    state.ui.invoiceSearchQuery = invoiceSearchInput?.value || "";
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderInvoicesPanel();
  };
  invoiceSearchInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    applyInvoiceSearch();
  });
  document.getElementById("applyInvoiceSearchButton")?.addEventListener("click", applyInvoiceSearch);

  document.getElementById("invoiceStatusFilter")?.addEventListener("change", (event) => {
    state.ui.invoiceStatusFilter = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoicePaymentFilter")?.addEventListener("change", (event) => {
    state.ui.invoicePaymentFilter = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoiceClientFilter")?.addEventListener("change", (event) => {
    state.ui.invoiceClientFilter = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoiceYearFilter")?.addEventListener("change", (event) => {
    state.ui.invoiceExportYear = event.currentTarget.value;
    state.ui.invoiceExportMonth = "all";
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoiceMonthFilter")?.addEventListener("change", (event) => {
    state.ui.invoiceExportMonth = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoiceLightDesignerFilter")?.addEventListener("change", (event) => {
    state.ui.invoiceLightDesignerFilter = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("invoiceSortMode")?.addEventListener("change", (event) => {
    state.ui.invoiceSortMode = event.currentTarget.value;
    state.ui.invoiceRegisterPage = 1;
    saveState(state);
    renderDashboard();
  });

  const paymentReconSearchInput = document.getElementById("paymentReconSearchInput");
  const applyPaymentReconSearch = () => {
    state.ui.paymentReconSearchQuery = paymentReconSearchInput?.value || "";
    state.ui.paymentReconPage = 1;
    saveState(state);
    renderInvoicesPanel();
  };
  paymentReconSearchInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    applyPaymentReconSearch();
  });
  document.getElementById("applyPaymentReconSearchButton")?.addEventListener("click", applyPaymentReconSearch);

  document.getElementById("paymentReconClientFilter")?.addEventListener("change", (event) => {
    state.ui.paymentReconClient = event.currentTarget.value;
    state.ui.paymentReconPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("paymentReconYearFilter")?.addEventListener("change", (event) => {
    state.ui.paymentReconYear = event.currentTarget.value;
    state.ui.paymentReconMonth = "all";
    state.ui.paymentReconPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("paymentReconMonthFilter")?.addEventListener("change", (event) => {
    state.ui.paymentReconMonth = event.currentTarget.value;
    state.ui.paymentReconPage = 1;
    saveState(state);
    renderDashboard();
  });

  wirePaginationControls(panel, "invoice-register", "invoiceRegisterPage", "invoiceRegisterPageSize", () => renderInvoicesPanel());
  wirePaginationControls(panel, "invoice-payments", "paymentReconPage", "paymentReconPageSize", () => renderInvoicesPanel());

  panel.querySelectorAll("[data-invoice-subtab]").forEach((button) => {
    button.addEventListener("click", () => {
      if ((button.dataset.invoiceSubtab || "create") !== activeInvoiceSubtab && !confirmDiscardDirtyForm("switch invoice sections")) {
        return;
      }
      clearDirtyForm();
      state.ui.invoiceSubtab = button.dataset.invoiceSubtab || "create";
      if (state.ui.invoiceSubtab === "create") {
        state.ui.markPaymentInvoiceId = null;
      }
      saveState(state);
      renderInvoicesPanel();
    });
  });

  document.getElementById("addInvoiceLineItem")?.addEventListener("click", () => {
    addLineItemRow();
    setDirtyForm("invoice");
  });
  form.elements.namedItem("amountPaid").addEventListener("input", updateTotalsPreview);
  wireDirtyFormTracking(form, "invoice");

  const draftClientNames = [...new Set(draftShows.map((show) => String(show.client || "").trim()).filter(Boolean))];
  const draftClientIds = [...new Set(draftShows.map((show) => String(show.clientId || getClientByName(show.client)?.id || "").trim()).filter(Boolean))];
  const templateInvoice = !editingInvoice && state.ui.invoiceDraftTemplate ? state.ui.invoiceDraftTemplate : null;
  const draftInvoice = !editingInvoice && !templateInvoice && draftShows.length ? {
    invoiceNumber: makeDefaultInvoiceNumber(),
    clientId: draftClientIds.length === 1 ? draftClientIds[0] : "",
    clientName: draftClientNames.length === 1 ? draftClientNames[0] : "",
    issueDate: dateKey(new Date()),
    dueDate: "",
    status: "draft",
    amountPaid: 0,
    notes: "",
    details: normalizeInvoiceDetails(),
    lineItems: draftShows.map((show) => ({
      id: uid("line"),
      showId: show.id,
      description: defaultInvoiceLineDescription(show),
      sac: state.ui.invoiceLineDefaults?.sac || "",
      discount: "",
      quantity: 1,
      unitRate: Number(show.amountShow || 0),
      amount: Number(show.amountShow || 0)
    }))
  } : null;
  const initialInvoice = editingInvoice || templateInvoice || draftInvoice || {
    invoiceNumber: makeDefaultInvoiceNumber(),
    clientId: "",
    clientName: "",
    issueDate: dateKey(new Date()),
    dueDate: "",
    status: "draft",
    amountPaid: 0,
    notes: "",
    details: normalizeInvoiceDetails(),
    lineItems: []
  };
  initialInvoice.details = normalizeInvoiceDetails(initialInvoice.details);
  initialInvoice.clientId = initialInvoice.clientId || getClientByName(initialInvoice.clientName)?.id || "";
  initialInvoice.dueDate = getDueDateFromTerms(initialInvoice.issueDate, initialInvoice.details.paymentTerms);
  const initialClient = getClientById(initialInvoice.clientId);
  if (initialClient) {
    if (!initialInvoice.details.clientGstin) {
      initialInvoice.details.clientGstin = initialClient.gstin || "";
    }
    if (!initialInvoice.details.clientBillingAddress) {
      initialInvoice.details.clientBillingAddress = initialClient.billingAddress || "";
    }
    if (!initialInvoice.details.placeOfSupply) {
      initialInvoice.details.placeOfSupply = initialClient.state || "";
    }
  }
  form.elements.namedItem("invoiceId").value = editingInvoice?.id || "";
  form.elements.namedItem("invoiceNumber").value = initialInvoice.invoiceNumber || "";
  form.elements.namedItem("clientId").value = initialInvoice.clientId || "";
  form.elements.namedItem("issueDate").value = initialInvoice.issueDate || dateKey(new Date());
  form.elements.namedItem("status").value = initialInvoice.status || "draft";
  form.elements.namedItem("amountPaid").value = initialInvoice.amountPaid ?? 0;
  form.elements.namedItem("notes").value = initialInvoice.notes || "";
  form.elements.namedItem("paymentTerms").value = initialInvoice.details.paymentTerms || "Net 15";
  syncCustomSelect(form.elements.namedItem("clientId"));
  lineItemsContainer.innerHTML = "";
  if (initialInvoice.lineItems.length) {
    initialInvoice.lineItems.forEach((item) => addLineItemRow(item));
  } else {
    addLineItemRow();
  }
  updateTotalsPreview();

  clientSelect.addEventListener("change", refreshLineItemShowOptions);

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const message = document.getElementById("invoiceFormMessage");
    message.textContent = "";
    const rows = [...lineItemsContainer.querySelectorAll("[data-invoice-line-item]")];
    const lineItems = rows.map((row, index) => {
      const quantity = Number(row.querySelector('input[name="lineQuantity"]').value || 0);
      const unitRate = Number(row.querySelector('input[name="lineUnitRate"]').value || 0);
      return {
        id: row.querySelector('input[name="lineId"]').value,
        showId: row.querySelector('input[name="lineShowId"]').value,
        description: row.querySelector('input[name="lineDescription"]').value.trim(),
        sac: row.querySelector('input[name="lineSac"]').value.trim(),
        customDetails: row.querySelector('input[name="lineCustomDetails"]').value.trim(),
        discount: row.querySelector('input[name="lineDiscount"]').value.trim(),
        quantity: quantity > 0 ? quantity : 1,
        unitRate,
        lineOrder: index
      };
    }).filter((item) => item.description);

    if (!lineItems.length) {
      message.textContent = "Add at least one line item.";
      return;
    }

    try {
      const selectedClient = getClientById(form.elements.namedItem("clientId").value);
      if (!selectedClient) {
        message.textContent = "Select a client from the client master.";
        return;
      }
      const issueDate = form.elements.namedItem("issueDate").value;
      const paymentTerms = form.elements.namedItem("paymentTerms").value.trim();
      const firstLineWithDefaults = lineItems.find((item) => item.description || item.sac);
      if (firstLineWithDefaults) {
        state.ui.invoiceLineDefaults = {
          description: firstLineWithDefaults.description || "",
          sac: firstLineWithDefaults.sac || ""
        };
      }
      const payload = await apiRequest("/api/admin/invoices", {
        method: "POST",
        body: JSON.stringify({
          id: form.elements.namedItem("invoiceId").value || undefined,
          invoiceNumber: form.elements.namedItem("invoiceNumber").value.trim(),
          clientId: selectedClient.id,
          clientName: selectedClient.name,
          issueDate,
          dueDate: getDueDateFromTerms(issueDate, paymentTerms),
          status: form.elements.namedItem("status").value,
          taxPercent: 0,
          amountPaid: Number(form.elements.namedItem("amountPaid").value || 0),
          notes: form.elements.namedItem("notes").value.trim(),
          details: {
            clientGstin: selectedClient.gstin || "",
            placeOfSupply: selectedClient.state || "",
            paymentTerms,
            clientBillingAddress: selectedClient.billingAddress || ""
          },
          lineItems
        })
      });
      applyServerState(payload);
      clearDirtyForm("invoice");
      resetInvoiceEditingState();
      state.ui.invoiceSubtab = "register";
      saveState(state);
      renderDashboard();
      showToast(editingInvoice ? "Invoice updated." : "Invoice saved.");
    } catch (error) {
      message.textContent = error.message;
    }
  });

  document.getElementById("deleteInvoiceButton")?.addEventListener("click", async () => {
    if (!editingInvoice) return;
    const confirmed = window.confirm(`Delete invoice "${editingInvoice.invoiceNumber}"?`);
    if (!confirmed) return;
    try {
      const payload = await apiRequest(`/api/admin/invoices/${encodeURIComponent(editingInvoice.id)}`, {
        method: "DELETE"
      });
      applyServerState(payload);
      resetInvoiceEditingState();
      state.ui.invoiceSubtab = "register";
      saveState(state);
      renderDashboard();
      showToast("Invoice deleted.");
    } catch (error) {
      showToast(error.message);
    }
  });

  panel.querySelectorAll("[data-edit-invoice]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!confirmDiscardDirtyForm("open another invoice")) return;
      clearDirtyForm();
      state.ui.editingInvoiceId = button.dataset.editInvoice;
      state.ui.invoiceDraftTemplate = null;
      state.ui.invoiceSubtab = "create";
      state.ui.markPaymentInvoiceId = null;
      saveState(state);
      renderDashboard();
    });
  });

  panel.querySelectorAll("[data-duplicate-invoice]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!confirmDiscardDirtyForm("duplicate this invoice")) return;
      const invoice = invoices.find((item) => item.id === button.dataset.duplicateInvoice);
      if (!invoice) return;
      clearDirtyForm();
      resetInvoiceEditingState();
      state.ui.invoiceDraftTemplate = buildInvoiceDraftTemplateFromInvoice(invoice);
      state.ui.invoiceSubtab = "create";
      state.ui.markPaymentInvoiceId = null;
      saveState(state);
      renderDashboard();
      showToast("Invoice duplicated into a new draft.");
    });
  });

  panel.querySelectorAll("[data-mark-payment]").forEach((button) => {
    button.addEventListener("click", () => {
      const invoiceId = button.dataset.markPayment;
      state.ui.markPaymentInvoiceId = state.ui.markPaymentInvoiceId === invoiceId ? null : invoiceId;
      state.ui.invoiceSubtab = "register";
      saveState(state);
      renderInvoicesPanel();
    });
  });

  panel.querySelectorAll("[data-cancel-payment]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.markPaymentInvoiceId = null;
      saveState(state);
      renderInvoicesPanel();
    });
  });

  panel.querySelectorAll("[data-payment-form]").forEach((paymentForm) => {
    paymentForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const invoiceId = paymentForm.dataset.paymentForm;
      const message = paymentForm.querySelector("[data-payment-message]");
      if (message) message.textContent = "";
      const formData = new FormData(paymentForm);
      try {
        const payload = await apiRequest(`/api/admin/invoices/${encodeURIComponent(invoiceId)}/payments`, {
          method: "POST",
          body: JSON.stringify({
            paymentDate: formData.get("paymentDate")?.toString() || "",
            amount: Number(formData.get("amount") || 0),
            note: formData.get("note")?.toString().trim() || ""
          })
        });
        applyServerState(payload);
        state.ui.markPaymentInvoiceId = null;
        saveState(state);
        renderDashboard();
        showToast("Payment recorded.");
      } catch (error) {
        if (message) {
          message.textContent = error.message;
        } else {
          showToast(error.message);
        }
      }
    });
  });

  panel.querySelectorAll("[data-preview-invoice]").forEach((button) => {
    button.addEventListener("click", () => {
      openInvoicePrintPreview(button.dataset.previewInvoice);
    });
  });

  panel.querySelectorAll("[data-open-payment-invoice]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!confirmDiscardDirtyForm("open this invoice")) return;
      clearDirtyForm();
      state.ui.editingInvoiceId = button.dataset.openPaymentInvoice;
      state.ui.invoiceDraftTemplate = null;
      state.ui.invoiceSubtab = "create";
      state.ui.markPaymentInvoiceId = null;
      saveState(state);
      renderDashboard();
    });
  });

  const invoicePreviewModal = document.getElementById("invoicePreviewModal");
  const closeInvoicePreviewButton = document.getElementById("closeInvoicePreview");
  const invoicePrintCopiesSelect = document.getElementById("invoicePrintCopies");
  if (closeInvoicePreviewButton) {
    closeInvoicePreviewButton.onclick = () => {
      closeInvoicePrintPreview();
    };
  }
  const printInvoicePreviewButton = document.getElementById("printInvoicePreview");
  if (printInvoicePreviewButton) {
    printInvoicePreviewButton.onclick = () => {
      const invoiceId = invoicePreviewModal?.dataset.invoiceId || "";
      if (invoiceId) {
        printInvoiceById(invoiceId, state.ui.invoicePrintCopies || 1);
      }
    };
  }
  if (invoicePrintCopiesSelect) {
    invoicePrintCopiesSelect.onchange = () => {
      state.ui.invoicePrintCopies = Number(invoicePrintCopiesSelect.value || 1);
      saveState(state);
      const invoiceId = invoicePreviewModal?.dataset.invoiceId || "";
      const previewContent = document.getElementById("invoicePreviewContent");
      const invoice = (state.invoices || []).find((item) => item.id === invoiceId);
      if (invoice && previewContent) {
        previewContent.innerHTML = getInvoiceDocumentMarkup(invoice, { copyCount: state.ui.invoicePrintCopies || 1 });
      }
    };
  }
  if (invoicePreviewModal) {
    invoicePreviewModal.onclick = (event) => {
      if (event.target === invoicePreviewModal) {
        closeInvoicePrintPreview();
      }
    };
  }
}

function renderGoogleEntriesPanel(user) {
  const panel = document.getElementById("googleEntriesPanel");
  const googleShows = sortShows(getGoogleLinkedShows(visibleShowsForUser(user)), "date");
  const currentMonthStart = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;
  const activeGoogleShows = googleShows.filter((show) => !show.googleArchived && (show.googlePinned || getShowEndDate(show) >= currentMonthStart));
  const archivedGoogleShows = googleShows.filter((show) => show.googleArchived || (!show.googlePinned && getShowEndDate(show) < currentMonthStart));
  const needsCompletion = activeGoogleShows.filter((show) => show.needsAdminCompletion);
  const synced = activeGoogleShows.filter((show) => !show.needsAdminCompletion);
  const archiveYearOptions = getArchiveYearOptions(archivedGoogleShows);
  const archiveMonthOptions = getArchiveMonthOptions(archivedGoogleShows, state.ui.googleArchiveYear || "all");
  const filteredArchivedShows = sortShows(filterArchivedGoogleShows(archivedGoogleShows), "date");
  const googleStatus = state.google || {};
  const googleMissingConfigText = Array.isArray(googleStatus.missingConfig) && googleStatus.missingConfig.length
    ? googleStatus.missingConfig.join(", ")
    : "";
  const googleEntriesView = state.ui.googleEntriesView || "needsCompletion";
  const googleEntriesMap = {
    needsCompletion: {
      title: "Needs Completion",
      empty: "No Google entries currently need admin completion.",
      shows: needsCompletion
    },
    synced: {
      title: "Synced",
      empty: "No synced Google entries in the current month window.",
      shows: synced
    },
    archived: {
      title: "Archived",
      empty: "No archived Google-linked entries for the selected month/year.",
      shows: filteredArchivedShows
    }
  };
  const selectedGoogleEntriesView = googleEntriesMap[googleEntriesView] || googleEntriesMap.needsCompletion;
  const googlePagination = getPaginationSlice(selectedGoogleEntriesView.shows, "googleEntriesPage", "googleEntriesPageSize");

  panel.innerHTML = `
    <div class="stack google-calendar-panel">
      <div class="form-header google-calendar-topbar">
        <div>
          <h3>Google Calendar</h3>
          <p class="muted-note">Use the tabs below to switch between entries that need admin completion, already synced entries, and archived Google-linked items.</p>
        </div>
        <div class="toolbar google-calendar-actions">
          <span class="pill">${activeGoogleShows.length} ${activeGoogleShows.length === 1 ? "entry" : "entries"}</span>
          ${googleStatus.connected
            ? `<button type="button" class="secondary small" id="googleSyncNowButton">Sync Now</button>`
            : `<button type="button" class="secondary small" id="googleConnectButton" ${googleStatus.configured ? "" : "disabled"}>${googleStatus.configured ? "Connect Google Calendar" : "Google Not Configured"}</button>`}
        </div>
      </div>
      ${!googleStatus.configured ? `
        <div class="assignment-card">
          <div class="meta"><strong>Google setup required:</strong> Set environment variables in Render and redeploy.</div>
          ${googleMissingConfigText ? `<div class="meta">Missing: ${googleMissingConfigText}</div>` : ""}
          ${googleStatus.redirectUri ? `<div class="meta">OAuth Redirect URI: ${googleStatus.redirectUri}</div>` : ""}
        </div>
      ` : ""}
      <div class="summary-grid google-calendar-summary-grid">
        <button type="button" class="summary-card summary-tab ${googleEntriesView === "needsCompletion" ? "is-active" : ""}" data-google-view="needsCompletion">
          <span class="summary-kicker">Needs Completion</span>
          <strong>${needsCompletion.length}</strong>
          <span class="summary-foot">Imported from Google and still missing admin-only details.</span>
        </button>
        <button type="button" class="summary-card summary-tab ${googleEntriesView === "synced" ? "is-active" : ""}" data-google-view="synced">
          <span class="summary-kicker">Synced</span>
          <strong>${synced.length}</strong>
          <span class="summary-foot">Linked entries already completed in PixelBug.</span>
        </button>
        <button type="button" class="summary-card summary-tab ${googleEntriesView === "archived" ? "is-active" : ""}" data-google-view="archived">
          <span class="summary-kicker">Archived</span>
          <strong>${archivedGoogleShows.length}</strong>
          <span class="summary-foot">Linked entries from older months or manually archived items.</span>
        </button>
      </div>
      ${(googleStatus.lastSyncAt || googleStatus.lastError) ? `
        <div class="assignment-card google-sync-meta-row google-calendar-meta-row">
          <div class="google-sync-meta-left">
            ${googleStatus.lastSyncAt ? `<div class="meta">Last Sync: ${new Date(googleStatus.lastSyncAt).toLocaleString("en-IN")}</div>` : ""}
            ${googleStatus.lastError ? `<div class="meta">Last Error: ${googleStatus.lastError}</div>` : ""}
          </div>
          <div class="google-sync-meta-right">
            <div class="meta"><strong>Status:</strong> ${googleStatus.connected ? "Connected" : googleStatus.configured ? "Ready to connect" : "Missing config"}</div>
            ${googleStatus.calendarId ? `<div class="meta">Calendar: ${googleStatus.calendarId}</div>` : ""}
          </div>
        </div>
      ` : `
        <div class="assignment-card google-sync-meta-row google-calendar-meta-row">
          <div class="google-sync-meta-left"></div>
          <div class="google-sync-meta-right">
            <div class="meta"><strong>Status:</strong> ${googleStatus.connected ? "Connected" : googleStatus.configured ? "Ready to connect" : "Missing config"}</div>
            ${googleStatus.calendarId ? `<div class="meta">Calendar: ${googleStatus.calendarId}</div>` : ""}
          </div>
        </div>
      `}
      <div class="form-header google-calendar-subhead" style="margin-top: 8px;">
        <div>
          <h4>${selectedGoogleEntriesView.title}</h4>
          <p class="muted-note">
            ${googleEntriesView === "archived"
              ? "Archived Google-linked entries can be filtered by month and year and restored whenever needed."
              : googleEntriesView === "synced"
                ? "Entries here are already linked and completed inside PixelBug."
                : "Entries here still need admin completion before they are fully ready in PixelBug."}
          </p>
        </div>
        <div class="toolbar google-calendar-subhead-actions">
          <span class="pill">${selectedGoogleEntriesView.shows.length} ${selectedGoogleEntriesView.shows.length === 1 ? "entry" : "entries"}</span>
        </div>
      </div>
      ${googleEntriesView === "archived" ? `
        <div class="toolbar google-calendar-archive-toolbar">
          <label class="field">
            <span>Year</span>
            <select id="googleArchiveYear">
              <option value="all" ${state.ui.googleArchiveYear === "all" ? "selected" : ""}>All Years</option>
              ${archiveYearOptions.map((year) => `<option value="${year}" ${state.ui.googleArchiveYear === year ? "selected" : ""}>${year}</option>`).join("")}
            </select>
          </label>
          <label class="field">
            <span>Month</span>
            <select id="googleArchiveMonth">
              <option value="all" ${state.ui.googleArchiveMonth === "all" ? "selected" : ""}>All Months</option>
              ${archiveMonthOptions.map((monthKey) => {
                const [year, month] = monthKey.split("-");
                return `<option value="${monthKey}" ${state.ui.googleArchiveMonth === monthKey ? "selected" : ""}>${monthLabel(Number(year), Number(month) - 1)}</option>`;
              }).join("")}
            </select>
          </label>
        </div>
      ` : ""}
      <div class="show-groups">
        ${googlePagination.items.length ? googlePagination.items.map((show) => {
          const { color } = getShowDisplayMeta(show, user);
          return `
          <article class="show-card google-linked-show-card" style="--google-entry-color:${escapeHtml(color || "#264653")}">
            <header>
              <div>
                <h4>${show.showName}</h4>
                <div class="meta">${formatDateRange(getShowStartDate(show), getShowEndDate(show))}</div>
              </div>
              <div class="toolbar">
                <span class="pill">${formatSyncStatus(show)}</span>
                ${googleEntriesView === "archived"
                  ? `<button type="button" class="secondary small" data-restore-google-show="${show.id}">Restore</button>`
                  : `<button type="button" class="secondary small" data-edit-show="${show.id}">${show.needsAdminCompletion ? "Complete Entry" : "Edit"}</button>
                     <button type="button" class="ghost small" data-archive-google-show="${show.id}">Archive</button>`}
              </div>
            </header>
            <div class="show-banner">
              <span class="show-banner-item">${show.location || "Location TBD"}</span>
              <span class="show-banner-item">${show.googleSyncSource === "google" ? "Imported from Google" : "Pushed from PixelBug"}</span>
              <span class="show-banner-item">${googleEntriesView === "archived"
                ? (show.googleArchivedAt ? `Archived ${new Date(show.googleArchivedAt).toLocaleDateString("en-IN")}` : "Older Month")
                : show.googleEventId ? "Linked" : "Unlinked"}</span>
            </div>
            <div class="stack" style="margin-top:12px;">
              <strong>Google Sync Fields</strong>
              <div class="assignment-card">
                <div class="meta">Crew Summary: ${(show.assignments || []).map((assignment) => getAssignmentCrewName(assignment)).filter(Boolean).join(", ") || "Unassigned"}</div>
                <div class="meta">Notes: ${show.googleNotes || "-"}</div>
                <div class="meta">Last Google Sync: ${show.googleLastSyncedAt ? new Date(show.googleLastSyncedAt).toLocaleString("en-IN") : "-"}</div>
                ${show.googleEventId ? `<div class="meta">Google Event ID: ${show.googleEventId}</div>` : ""}
              </div>
            </div>
          </article>
        `;
        }).join("") : `<p>${selectedGoogleEntriesView.empty}</p>`}
        </div>
        ${renderPaginationControls("google-entries", googlePagination, "entries")}
    </div>
  `;

  enhanceCustomSelects(panel);

  panel.querySelectorAll("[data-google-view]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.googleEntriesView = button.dataset.googleView;
      state.ui.googleEntriesPage = 1;
      saveState(state);
      renderDashboard();
    });
  });

  wirePaginationControls(panel, "google-entries", "googleEntriesPage", "googleEntriesPageSize", () => renderGoogleEntriesPanel(user));

  panel.querySelectorAll("[data-edit-show]").forEach((button) => {
    button.addEventListener("click", () => fillShowForm(button.dataset.editShow));
  });

  panel.querySelectorAll("[data-archive-google-show]").forEach((button) => {
    button.addEventListener("click", async () => {
      const show = state.shows.find((item) => item.id === button.dataset.archiveGoogleShow);
      if (!show) return;
      show.googleArchived = true;
      show.googleArchivedAt = new Date().toISOString();
      show.googlePinned = false;
      try {
        await syncAdminState();
        saveState(state);
        render();
        showToast("Google entry archived.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        render();
      }
    });
  });

  panel.querySelectorAll("[data-restore-google-show]").forEach((button) => {
    button.addEventListener("click", async () => {
      const show = state.shows.find((item) => item.id === button.dataset.restoreGoogleShow);
      if (!show) return;
      show.googleArchived = false;
      show.googleArchivedAt = "";
      show.googlePinned = true;
      try {
        await syncAdminState();
        saveState(state);
        render();
        showToast("Google entry restored.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        render();
      }
    });
  });

  document.getElementById("googleArchiveYear")?.addEventListener("change", (event) => {
    state.ui.googleArchiveYear = event.currentTarget.value;
    state.ui.googleEntriesPage = 1;
    if (state.ui.googleArchiveYear !== "all") {
      const monthOptions = getArchiveMonthOptions(archivedGoogleShows, state.ui.googleArchiveYear);
      if (state.ui.googleArchiveMonth !== "all" && !monthOptions.includes(state.ui.googleArchiveMonth)) {
        state.ui.googleArchiveMonth = "all";
      }
    }
    saveState(state);
    render();
  });

  document.getElementById("googleArchiveMonth")?.addEventListener("change", (event) => {
    state.ui.googleArchiveMonth = event.currentTarget.value;
    state.ui.googleEntriesPage = 1;
    saveState(state);
    render();
  });

  document.getElementById("googleConnectButton")?.addEventListener("click", async () => {
    try {
      const payload = await apiRequest("/api/admin/google/auth");
      window.location.href = payload.authUrl;
    } catch (error) {
      showToast(error.message);
    }
  });

  document.getElementById("googleSyncNowButton")?.addEventListener("click", async () => {
    try {
      const payload = await apiRequest("/api/admin/google/sync", { method: "POST" });
      applyServerState(payload);
      saveState(state);
      render();
      showToast("Google Calendar synced.");
    } catch (error) {
      showToast(error.message);
    }
  });
}

function renderPayoutsPanel() {
  const panel = document.getElementById("payoutsPanel");
  const payoutRows = getCrewPayoutRows();
  const payoutYearOptions = getCrewPayoutYearOptions(payoutRows);
  if (!["all", ...payoutYearOptions].includes(state.ui.payoutYear)) {
    state.ui.payoutYear = "all";
  }
  const payoutMonthKeys = getCrewPayoutMonthOptions(payoutRows, state.ui.payoutYear || "all");
  if (state.ui.payoutMonth !== "all" && !payoutMonthKeys.includes(state.ui.payoutMonth)) {
    state.ui.payoutMonth = "all";
  }
  const payoutCrewOptions = ["all", ...new Set(payoutRows.map((row) => row.crewName).filter(Boolean))].sort((a, b) => {
    if (a === "all") return -1;
    if (b === "all") return 1;
    return a.localeCompare(b);
  });
  if (!payoutCrewOptions.includes(state.ui.payoutCrew)) {
    state.ui.payoutCrew = "all";
  }
  const payoutClientOptions = ["all", ...new Set(payoutRows.map((row) => row.clientLabel).filter(Boolean))].sort((a, b) => {
    if (a === "all") return -1;
    if (b === "all") return 1;
    return a.localeCompare(b);
  });
  if (!payoutClientOptions.includes(state.ui.payoutClient)) {
    state.ui.payoutClient = "all";
  }

  const filteredRows = filterCrewPayoutRows(payoutRows);
  const payoutPagination = getPaginationSlice(filteredRows, "payoutPage", "payoutPageSize");
  const summary = {
    totalPayout: filteredRows.reduce((sum, row) => sum + Number(row.operatorAmount || 0), 0),
    assignments: filteredRows.length,
    shows: new Set(filteredRows.map((row) => row.showId).filter(Boolean)).size,
    crew: new Set(filteredRows.map((row) => row.crewName).filter(Boolean)).size
  };

  panel.innerHTML = `
    <div class="stack">
      <div class="form-header">
        <div>
          <h3>Crew Payouts</h3>
          <p class="muted-note">Track operator payouts from show assignments and export month-wise payout sheets.</p>
        </div>
        <div class="toolbar">
          <span class="pill">${summary.assignments} ${summary.assignments === 1 ? "assignment" : "assignments"}</span>
        </div>
      </div>
      <div class="summary-grid">
        <div class="summary-card">
          <span class="summary-kicker">Total Payout</span>
          <strong>${formatCurrency(summary.totalPayout)}</strong>
          <span class="summary-foot">Operator payouts in this view.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Assignments</span>
          <strong>${summary.assignments}</strong>
          <span class="summary-foot">Payable crew rows.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Shows</span>
          <strong>${summary.shows}</strong>
          <span class="summary-foot">Shows covered by these payouts.</span>
        </div>
        <div class="summary-card">
          <span class="summary-kicker">Crew</span>
          <strong>${summary.crew}</strong>
          <span class="summary-foot">Crew members in this view.</span>
        </div>
      </div>
      <div class="shows-toolbar">
        <div class="shows-toolbar-top">
          <label class="sort-control invoice-search-control">
            <span>Search</span>
            <input type="search" id="payoutSearchInput" placeholder="Show, client, crew, location" value="${escapeHtml(state.ui.payoutSearchQuery || "")}" autocomplete="off">
          </label>
          <label class="sort-control">
            <span>Crew</span>
            <select id="payoutCrewFilter">
              ${payoutCrewOptions.map((option) => `<option value="${option}" ${state.ui.payoutCrew === option ? "selected" : ""}>${option === "all" ? "All Crew" : escapeHtml(option)}</option>`).join("")}
            </select>
          </label>
          <label class="sort-control">
            <span>Client</span>
            <select id="payoutClientFilter">
              ${payoutClientOptions.map((option) => `<option value="${option}" ${state.ui.payoutClient === option ? "selected" : ""}>${option === "all" ? "All Clients" : escapeHtml(option)}</option>`).join("")}
            </select>
          </label>
        </div>
        <div class="shows-toolbar-top invoice-toolbar-secondary">
          <label class="sort-control">
            <span>Year</span>
            <select id="payoutYearFilter">
              <option value="all" ${state.ui.payoutYear === "all" ? "selected" : ""}>All Years</option>
              ${payoutYearOptions.map((year) => `<option value="${year}" ${state.ui.payoutYear === year ? "selected" : ""}>${year}</option>`).join("")}
            </select>
          </label>
          <label class="sort-control">
            <span>Month</span>
            <select id="payoutMonthFilter">
              <option value="all" ${state.ui.payoutMonth === "all" ? "selected" : ""}>All Months</option>
              ${payoutMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.payoutMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
            </select>
          </label>
          <button type="button" class="secondary" id="exportPayoutsButton" ${filteredRows.length ? "" : "disabled"}>Export Excel</button>
        </div>
      </div>
      <div class="client-ledger-table-wrap">
        <table class="client-ledger-table">
          <thead>
            <tr>
              <th>Show Date</th>
              <th>Show</th>
              <th>Client</th>
              <th>Crew</th>
              <th>Light Designer</th>
              <th>Location</th>
              <th>Amount</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${payoutPagination.items.length ? payoutPagination.items.map((row) => `
              <tr>
                <td>${escapeHtml(row.showDateLabel)}</td>
                <td><strong>${escapeHtml(row.showName)}</strong>${row.notes ? `<div class="meta">${escapeHtml(row.notes)}</div>` : ""}</td>
                <td>${escapeHtml(row.clientLabel || "-")}</td>
                <td>${escapeHtml(row.crewName)}</td>
                <td>${escapeHtml(row.lightDesigner || "-")}</td>
                <td>${escapeHtml(row.location || "-")}</td>
                <td><strong>${formatCurrency(row.operatorAmount)}</strong></td>
                <td><span class="pill ${row.status === "tentative" ? "pill-warning" : ""}">${escapeHtml(row.statusLabel)}</span></td>
              </tr>
            `).join("") : `<tr><td colspan="8">No payout rows match the current filters.</td></tr>`}
          </tbody>
        </table>
      </div>
      ${renderPaginationControls("payouts", payoutPagination, "payout rows")}
    </div>
  `;

  document.getElementById("payoutSearchInput")?.addEventListener("input", (event) => {
    state.ui.payoutSearchQuery = event.currentTarget.value;
    state.ui.payoutPage = 1;
    saveState(state);
    const cursorPosition = event.currentTarget.selectionStart;
    renderPayoutsPanel();
    const searchInput = document.getElementById("payoutSearchInput");
    searchInput?.focus();
    if (searchInput && cursorPosition !== null) {
      searchInput.setSelectionRange(cursorPosition, cursorPosition);
    }
  });

  document.getElementById("payoutCrewFilter")?.addEventListener("change", (event) => {
    state.ui.payoutCrew = event.currentTarget.value;
    state.ui.payoutPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("payoutClientFilter")?.addEventListener("change", (event) => {
    state.ui.payoutClient = event.currentTarget.value;
    state.ui.payoutPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("payoutYearFilter")?.addEventListener("change", (event) => {
    state.ui.payoutYear = event.currentTarget.value;
    state.ui.payoutMonth = "all";
    state.ui.payoutPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("payoutMonthFilter")?.addEventListener("change", (event) => {
    state.ui.payoutMonth = event.currentTarget.value;
    state.ui.payoutPage = 1;
    saveState(state);
    renderDashboard();
  });

  document.getElementById("exportPayoutsButton")?.addEventListener("click", () => {
    if (!filteredRows.length) {
      showToast("No payout rows to export.");
      return;
    }
    const exportKey = [
      state.ui.payoutYear !== "all" ? state.ui.payoutYear : "all",
      state.ui.payoutMonth !== "all" ? state.ui.payoutMonth : "",
      state.ui.payoutCrew !== "all" ? safeFileNamePart(state.ui.payoutCrew) : "",
      state.ui.payoutClient !== "all" ? safeFileNamePart(state.ui.payoutClient) : ""
    ].filter(Boolean).join("-");
    exportCrewPayoutsExcel(filteredRows, exportKey);
    showToast("Crew payout export ready.");
  });

  wirePaginationControls(panel, "payouts", "payoutPage", "payoutPageSize", () => renderPayoutsPanel());
}

function renderDocumentCenterPanel() {
  const panel = document.getElementById("documentsPanel");
  const user = getCurrentUser();
  const documentShowsBase = visibleShowsForUser(user);
  const documentShowYearOptions = [...new Set(documentShowsBase.map((show) => String(getShowStartDate(show) || "").slice(0, 4)).filter(Boolean))].sort((a, b) => b.localeCompare(a));
  if (!["all", ...documentShowYearOptions].includes(state.ui.documentShowsYear)) {
    state.ui.documentShowsYear = "all";
  }
  const documentShowMonthKeys = [...new Set(documentShowsBase
    .filter((show) => state.ui.documentShowsYear === "all" || String(getShowStartDate(show) || "").slice(0, 4) === state.ui.documentShowsYear)
    .map((show) => String(getShowStartDate(show) || "").slice(0, 7))
    .filter(Boolean))].sort((a, b) => b.localeCompare(a));
  if (state.ui.documentShowsMonth !== "all" && !documentShowMonthKeys.includes(state.ui.documentShowsMonth)) {
    state.ui.documentShowsMonth = "all";
  }
  const filteredShows = sortShows(documentShowsBase.filter((show) => {
    const showYear = String(getShowStartDate(show) || "").slice(0, 4);
    const showMonth = String(getShowStartDate(show) || "").slice(0, 7);
    const matchesYear = state.ui.documentShowsYear === "all" || showYear === state.ui.documentShowsYear;
    const matchesMonth = state.ui.documentShowsMonth === "all" || showMonth === state.ui.documentShowsMonth;
    return matchesYear && matchesMonth;
  }), "date");

  const allInvoicesSorted = sortInvoices(state.invoices || []);
  const documentInvoiceYearOptions = getInvoiceYearOptions(allInvoicesSorted);
  if (!["all", ...documentInvoiceYearOptions].includes(state.ui.documentInvoicesYear)) {
    state.ui.documentInvoicesYear = "all";
  }
  const documentInvoiceMonthKeys = getInvoiceMonthOptions(allInvoicesSorted, state.ui.documentInvoicesYear || "all");
  if (state.ui.documentInvoicesMonth !== "all" && !documentInvoiceMonthKeys.includes(state.ui.documentInvoicesMonth)) {
    state.ui.documentInvoicesMonth = "all";
  }
  const documentInvoiceClientOptions = ["all", ...getInvoiceClientOptions(allInvoicesSorted)];
  if (!documentInvoiceClientOptions.includes(state.ui.documentInvoicesClient)) {
    state.ui.documentInvoicesClient = "all";
  }
  const filteredInvoices = allInvoicesSorted.filter((invoice) => {
    const year = String(invoice.issueDate || "").slice(0, 4);
    const month = String(invoice.issueDate || "").slice(0, 7);
    const client = String(invoice.clientName || "").trim();
    const matchesYear = state.ui.documentInvoicesYear === "all" || year === state.ui.documentInvoicesYear;
    const matchesMonth = state.ui.documentInvoicesMonth === "all" || month === state.ui.documentInvoicesMonth;
    const matchesClient = state.ui.documentInvoicesClient === "all" || client === state.ui.documentInvoicesClient;
    return matchesYear && matchesMonth && matchesClient;
  });

  const allClients = getSortedClients();
  const filteredClients = allClients.filter((client) => {
    if (state.ui.documentClientsGstFilter === "missing") {
      return !String(client.gstin || "").trim();
    }
    if (state.ui.documentClientsGstFilter === "available") {
      return Boolean(String(client.gstin || "").trim());
    }
    return true;
  });

  const documentClientFinancialYearOptions = getClientInvoiceYearOptions();
  if (!["all", ...documentClientFinancialYearOptions].includes(state.ui.documentClientFinancialYear)) {
    state.ui.documentClientFinancialYear = "all";
  }
  const documentClientFinancialMonthKeys = getClientInvoiceMonthOptions(state.ui.documentClientFinancialYear || "all");
  if (state.ui.documentClientFinancialMonth !== "all" && !documentClientFinancialMonthKeys.includes(state.ui.documentClientFinancialMonth)) {
    state.ui.documentClientFinancialMonth = "all";
  }
  const documentClientFinancialClientOptions = [{ value: "all", label: "All Clients" }, ...allClients.map((client) => ({ value: client.id, label: getClientDisplayName(client) }))];
  if (!documentClientFinancialClientOptions.some((option) => option.value === state.ui.documentClientFinancialClient)) {
    state.ui.documentClientFinancialClient = "all";
  }
  const documentLedgerClientOptions = [{ value: "all", label: "Select client" }, ...allClients.map((client) => ({ value: client.id, label: getClientDisplayName(client) }))];
  if (!documentLedgerClientOptions.some((option) => option.value === state.ui.documentLedgerClient)) {
    state.ui.documentLedgerClient = "all";
  }
  const selectedClientIds = state.ui.documentClientFinancialClient !== "all"
    ? [state.ui.documentClientFinancialClient]
    : allClients.map((client) => client.id);
  const clientFinancialRows = getClientExportRows(selectedClientIds, state.ui.documentClientFinancialYear || "all", state.ui.documentClientFinancialMonth || "all");
  const selectedLedgerClient = state.ui.documentLedgerClient !== "all" ? getClientById(state.ui.documentLedgerClient) : null;
  const selectedLedgerEntries = selectedLedgerClient ? getClientLedgerEntries(selectedLedgerClient.id) : [];

  const allPayoutRows = getCrewPayoutRows();
  const documentPayoutYearOptions = getCrewPayoutYearOptions(allPayoutRows);
  if (!["all", ...documentPayoutYearOptions].includes(state.ui.documentPayoutYear)) {
    state.ui.documentPayoutYear = "all";
  }
  const documentPayoutMonthKeys = getCrewPayoutMonthOptions(allPayoutRows, state.ui.documentPayoutYear || "all");
  if (state.ui.documentPayoutMonth !== "all" && !documentPayoutMonthKeys.includes(state.ui.documentPayoutMonth)) {
    state.ui.documentPayoutMonth = "all";
  }
  const documentPayoutCrewOptions = ["all", ...new Set(allPayoutRows.map((row) => row.crewName).filter(Boolean))].sort((a, b) => {
    if (a === "all") return -1;
    if (b === "all") return 1;
    return a.localeCompare(b);
  });
  if (!documentPayoutCrewOptions.includes(state.ui.documentPayoutCrew)) {
    state.ui.documentPayoutCrew = "all";
  }
  const documentPayoutClientOptions = ["all", ...new Set(allPayoutRows.map((row) => row.clientLabel).filter(Boolean))].sort((a, b) => {
    if (a === "all") return -1;
    if (b === "all") return 1;
    return a.localeCompare(b);
  });
  if (!documentPayoutClientOptions.includes(state.ui.documentPayoutClient)) {
    state.ui.documentPayoutClient = "all";
  }
  const filteredPayoutRows = allPayoutRows.filter((row) => {
    const year = String(row.showDate || "").slice(0, 4);
    const month = String(row.showDate || "").slice(0, 7);
    const matchesYear = state.ui.documentPayoutYear === "all" || year === state.ui.documentPayoutYear;
    const matchesMonth = state.ui.documentPayoutMonth === "all" || month === state.ui.documentPayoutMonth;
    const matchesCrew = state.ui.documentPayoutCrew === "all" || row.crewName === state.ui.documentPayoutCrew;
    const matchesClient = state.ui.documentPayoutClient === "all" || row.clientLabel === state.ui.documentPayoutClient;
    return matchesYear && matchesMonth && matchesCrew && matchesClient;
  });

  panel.innerHTML = `
    <div class="stack">
      <div class="form-header">
        <div>
          <h3>Document Center</h3>
          <p class="muted-note">Generate the main exports and ledgers from one place using the filters you already set in the app.</p>
        </div>
      </div>
      <div class="client-detail-grid document-center-grid">
        <section class="assignment-card">
          <strong>Shows Export</strong>
          <p class="meta">Choose the show period here and export directly.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Year</span>
              <select id="documentShowsYear">
                <option value="all" ${state.ui.documentShowsYear === "all" ? "selected" : ""}>All Years</option>
                ${documentShowYearOptions.map((year) => `<option value="${year}" ${state.ui.documentShowsYear === year ? "selected" : ""}>${year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="documentShowsMonth">
                <option value="all" ${state.ui.documentShowsMonth === "all" ? "selected" : ""}>All Months</option>
                ${documentShowMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.documentShowsMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${filteredShows.length} rows</span>
            <button type="button" class="secondary" id="documentExportShows" ${filteredShows.length ? "" : "disabled"}>Export Shows</button>
          </div>
        </section>
        <section class="assignment-card">
          <strong>Invoice Register Export</strong>
          <p class="meta">Filter invoice export here without leaving the document hub.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Year</span>
              <select id="documentInvoicesYear">
                <option value="all" ${state.ui.documentInvoicesYear === "all" ? "selected" : ""}>All Years</option>
                ${documentInvoiceYearOptions.map((year) => `<option value="${year}" ${state.ui.documentInvoicesYear === year ? "selected" : ""}>${year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="documentInvoicesMonth">
                <option value="all" ${state.ui.documentInvoicesMonth === "all" ? "selected" : ""}>All Months</option>
                ${documentInvoiceMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.documentInvoicesMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Client</span>
              <select id="documentInvoicesClient">
                ${documentInvoiceClientOptions.map((option) => `<option value="${option}" ${state.ui.documentInvoicesClient === option ? "selected" : ""}>${option === "all" ? "All Clients" : escapeHtml(option)}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${filteredInvoices.length} rows</span>
            <button type="button" class="secondary" id="documentExportInvoices" ${filteredInvoices.length ? "" : "disabled"}>Export Invoices</button>
          </div>
        </section>
        <section class="assignment-card">
          <strong>Client Financial Export</strong>
          <p class="meta">Pick the exact client finance slice you want to export.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Year</span>
              <select id="documentClientFinancialYear">
                <option value="all" ${state.ui.documentClientFinancialYear === "all" ? "selected" : ""}>All Years</option>
                ${documentClientFinancialYearOptions.map((year) => `<option value="${year}" ${state.ui.documentClientFinancialYear === year ? "selected" : ""}>${year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="documentClientFinancialMonth">
                <option value="all" ${state.ui.documentClientFinancialMonth === "all" ? "selected" : ""}>All Months</option>
                ${documentClientFinancialMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.documentClientFinancialMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Client</span>
              <select id="documentClientFinancialClient">
                ${documentClientFinancialClientOptions.map((option) => `<option value="${option.value}" ${state.ui.documentClientFinancialClient === option.value ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${clientFinancialRows.length} rows</span>
            <button type="button" class="secondary" id="documentExportClientFinancials" ${clientFinancialRows.length ? "" : "disabled"}>Export Client Financials</button>
          </div>
        </section>
        <section class="assignment-card">
          <strong>Client Ledger</strong>
          <p class="meta">Pick a client and export or print the running ledger directly from here.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Client</span>
              <select id="documentLedgerClient">
                ${documentLedgerClientOptions.map((option) => `<option value="${option.value}" ${state.ui.documentLedgerClient === option.value ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${selectedLedgerClient ? `${selectedLedgerEntries.length} entries` : "Pick a client"}</span>
            <button type="button" class="secondary" id="documentExportClientLedger" ${(selectedLedgerClient && selectedLedgerEntries.length) ? "" : "disabled"}>Export Ledger</button>
            <button type="button" class="ghost" id="documentPrintClientLedger" ${(selectedLedgerClient && selectedLedgerEntries.length) ? "" : "disabled"}>Print Ledger</button>
          </div>
        </section>
        <section class="assignment-card">
          <strong>Client Master Export</strong>
          <p class="meta">Control GST completeness directly from here.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>GST Details</span>
              <select id="documentClientsGstFilter">
                <option value="all" ${state.ui.documentClientsGstFilter === "all" ? "selected" : ""}>All Clients</option>
                <option value="missing" ${state.ui.documentClientsGstFilter === "missing" ? "selected" : ""}>Empty GST Details</option>
                <option value="available" ${state.ui.documentClientsGstFilter === "available" ? "selected" : ""}>With GST Details</option>
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${filteredClients.length} rows</span>
            <button type="button" class="secondary" id="documentExportClients" ${filteredClients.length ? "" : "disabled"}>Export Clients</button>
          </div>
        </section>
        <section class="assignment-card">
          <strong>Crew Payout Export</strong>
          <p class="meta">Choose the payout period, crew, and client directly here.</p>
          <div class="shows-toolbar-top">
            <label class="sort-control">
              <span>Year</span>
              <select id="documentPayoutYear">
                <option value="all" ${state.ui.documentPayoutYear === "all" ? "selected" : ""}>All Years</option>
                ${documentPayoutYearOptions.map((year) => `<option value="${year}" ${state.ui.documentPayoutYear === year ? "selected" : ""}>${year}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Month</span>
              <select id="documentPayoutMonth">
                <option value="all" ${state.ui.documentPayoutMonth === "all" ? "selected" : ""}>All Months</option>
                ${documentPayoutMonthKeys.map((monthKey) => `<option value="${monthKey}" ${state.ui.documentPayoutMonth === monthKey ? "selected" : ""}>${monthGroupLabel(`${monthKey}-01`).split(" ")[0]}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Crew</span>
              <select id="documentPayoutCrew">
                ${documentPayoutCrewOptions.map((option) => `<option value="${option}" ${state.ui.documentPayoutCrew === option ? "selected" : ""}>${option === "all" ? "All Crew" : escapeHtml(option)}</option>`).join("")}
              </select>
            </label>
            <label class="sort-control">
              <span>Client</span>
              <select id="documentPayoutClient">
                ${documentPayoutClientOptions.map((option) => `<option value="${option}" ${state.ui.documentPayoutClient === option ? "selected" : ""}>${option === "all" ? "All Clients" : escapeHtml(option)}</option>`).join("")}
              </select>
            </label>
          </div>
          <div class="toolbar">
            <span class="pill">${filteredPayoutRows.length} rows</span>
            <button type="button" class="secondary" id="documentExportPayouts" ${filteredPayoutRows.length ? "" : "disabled"}>Export Payouts</button>
          </div>
        </section>
      </div>
    </div>
  `;

  document.getElementById("documentExportShows")?.addEventListener("click", () => {
    if (!filteredShows.length) {
      showToast("No shows to export.");
      return;
    }
    const exportKey = state.ui.documentShowsMonth !== "all"
      ? state.ui.documentShowsMonth
      : state.ui.documentShowsYear !== "all"
        ? state.ui.documentShowsYear
        : "all-shows";
    downloadSingleSheetWorkbook(exportKey, buildShowsSheetXml(filteredShows), `pixelbug-${safeFileNamePart(exportKey)}-shows.xlsx`);
    showToast("Shows export ready.");
  });

  document.getElementById("documentExportInvoices")?.addEventListener("click", () => {
    if (!filteredInvoices.length) {
      showToast("No invoices to export.");
      return;
    }
    const exportKey = state.ui.documentInvoicesMonth !== "all"
      ? state.ui.documentInvoicesMonth
      : state.ui.documentInvoicesYear !== "all"
        ? state.ui.documentInvoicesYear
        : "all";
    exportInvoicesExcel(exportKey, filteredInvoices, getInvoiceExportColumns(), {
      splitByMonth: state.ui.documentInvoicesMonth === "all"
    });
    showToast("Invoice export ready.");
  });

  document.getElementById("documentExportClientFinancials")?.addEventListener("click", () => {
    if (!clientFinancialRows.length) {
      showToast("No client financial rows to export.");
      return;
    }
    const exportKeyParts = [
      state.ui.documentClientFinancialYear !== "all" ? state.ui.documentClientFinancialYear : "",
      state.ui.documentClientFinancialMonth !== "all" ? state.ui.documentClientFinancialMonth : "",
      state.ui.documentClientFinancialClient !== "all" ? safeFileNamePart(getClientDisplayValue(state.ui.documentClientFinancialClient)) : ""
    ].filter(Boolean);
    const exportKey = exportKeyParts.join("-") || "all";
    downloadSingleSheetWorkbook("Client Invoices", buildClientInvoiceExportSheetXml(clientFinancialRows), `pixelbug-${safeFileNamePart(exportKey)}-client-invoices.xlsx`);
    showToast("Client financial export ready.");
  });

  document.getElementById("documentExportClientLedger")?.addEventListener("click", () => {
    if (!selectedLedgerClient || !selectedLedgerEntries.length) {
      showToast("Select a client with ledger entries first.");
      return;
    }
    exportClientLedgerExcel(selectedLedgerClient);
    showToast("Client ledger export ready.");
  });

  document.getElementById("documentPrintClientLedger")?.addEventListener("click", () => {
    if (!selectedLedgerClient || !selectedLedgerEntries.length) {
      showToast("Select a client with ledger entries first.");
      return;
    }
    printClientLedger(selectedLedgerClient.id);
  });

  document.getElementById("documentExportClients")?.addEventListener("click", () => {
    if (!filteredClients.length) {
      showToast("No clients to export.");
      return;
    }
    downloadSingleSheetWorkbook("Clients", buildClientsSheetXml(filteredClients), "pixelbug-clients.xlsx");
    showToast("Client master export ready.");
  });

  document.getElementById("documentExportPayouts")?.addEventListener("click", () => {
    if (!filteredPayoutRows.length) {
      showToast("No payout rows to export.");
      return;
    }
    const exportKey = [
      state.ui.documentPayoutYear !== "all" ? state.ui.documentPayoutYear : "all",
      state.ui.documentPayoutMonth !== "all" ? state.ui.documentPayoutMonth : "",
      state.ui.documentPayoutCrew !== "all" ? safeFileNamePart(state.ui.documentPayoutCrew) : "",
      state.ui.documentPayoutClient !== "all" ? safeFileNamePart(state.ui.documentPayoutClient) : ""
    ].filter(Boolean).join("-");
    exportCrewPayoutsExcel(filteredPayoutRows, exportKey);
    showToast("Crew payout export ready.");
  });

  [
    "documentShowsYear",
    "documentShowsMonth",
    "documentInvoicesYear",
    "documentInvoicesMonth",
    "documentInvoicesClient",
    "documentClientFinancialYear",
    "documentClientFinancialMonth",
    "documentClientFinancialClient",
    "documentLedgerClient",
    "documentClientsGstFilter",
    "documentPayoutYear",
    "documentPayoutMonth",
    "documentPayoutCrew",
    "documentPayoutClient"
  ].forEach((fieldId) => {
    document.getElementById(fieldId)?.addEventListener("change", (event) => {
      state.ui[fieldId] = event.currentTarget.value;
      saveState(state);
      renderDocumentCenterPanel();
    });
  });
}

function renderClientsPanel() {
  const panel = document.getElementById("clientsPanel");
  const clients = getSortedClients();
  const clientSearchQuery = String(state.ui.clientSearchQuery || "").trim().toLowerCase();
  const clientGstFilter = state.ui.clientGstFilter || "all";
  const filteredClients = clients.filter((client) => {
    const matchesSearch = !clientSearchQuery || [
      getClientDisplayName(client),
      client.name,
      client.state,
      client.gstin,
      client.contactName,
      client.contactEmail,
      client.contactPhone,
      client.billingAddress,
      client.notes
    ].some((value) => String(value || "").toLowerCase().includes(clientSearchQuery));
    if (!matchesSearch) return false;
    if (clientGstFilter === "missing") {
      return !String(client.gstin || "").trim();
    }
    if (clientGstFilter === "available") {
      return Boolean(String(client.gstin || "").trim());
    }
    return true;
  });
  const clientPagination = getPaginationSlice(filteredClients, "clientsPage", "clientsPageSize");
  const activeClientsSubtab = state.ui.clientsSubtab || "list";
  if (state.ui.selectedClientDetailId && !clients.some((client) => client.id === state.ui.selectedClientDetailId)) {
    state.ui.selectedClientDetailId = null;
  }
  const showsByClientId = new Map();
  const invoicesByClientId = new Map();
  state.shows.forEach((show) => {
    const key = show.clientId || getClientByName(show.client)?.id || "";
    if (!key) return;
    showsByClientId.set(key, (showsByClientId.get(key) || 0) + 1);
  });
  state.invoices.forEach((invoice) => {
    const key = invoice.clientId || getClientByName(invoice.clientName)?.id || "";
    if (!key) return;
    invoicesByClientId.set(key, (invoicesByClientId.get(key) || 0) + 1);
  });
  const selectedClientDetail = state.ui.selectedClientDetailId ? getClientById(state.ui.selectedClientDetailId) : null;
  const selectedClientShows = selectedClientDetail
    ? sortShows(state.shows.filter((show) => (show.clientId || getClientByName(show.client)?.id || "") === selectedClientDetail.id), "date")
    : [];
  const selectedClientInvoices = selectedClientDetail
    ? sortInvoices((state.invoices || []).filter((invoice) => (invoice.clientId || getClientByName(invoice.clientName)?.id || "") === selectedClientDetail.id))
    : [];
  const selectedClientPayments = selectedClientInvoices.flatMap((invoice) =>
    (Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : []).map((payment) => ({
      ...payment,
      invoiceNumber: invoice.invoiceNumber
    }))
  ).sort((a, b) => String(b.paymentDate || "").localeCompare(String(a.paymentDate || "")));
  const selectedClientLedgerEntries = selectedClientDetail ? getClientLedgerEntries(selectedClientDetail.id) : [];
  const clientOutstanding = selectedClientInvoices.reduce((sum, invoice) => sum + Number(invoice.balanceDue || 0), 0);
  const clientPaid = selectedClientInvoices.reduce((sum, invoice) => sum + Number(invoice.amountPaid || 0), 0);
  const clientBilled = selectedClientInvoices.reduce((sum, invoice) => sum + Number(invoice.totalAmount || 0), 0);

  panel.innerHTML = `
    <div class="stack">
      <div class="form-header">
        <div>
          <h3>Clients</h3>
          <p class="muted-note">Maintain each client once, then reuse it in shows and invoices.</p>
        </div>
        <span class="pill">${clients.length} ${clients.length === 1 ? "client" : "clients"}</span>
      </div>
      <div class="clients-subtab-row">
        <div class="invoice-subtabs" role="tablist" aria-label="Client sections">
          <button type="button" class="${activeClientsSubtab === "create" ? "is-active" : ""}" data-clients-subtab="create">${getDirtyTabLabel("Create Clients", "client")}</button>
          <button type="button" class="${activeClientsSubtab === "list" ? "is-active" : ""}" data-clients-subtab="list">Clients</button>
        </div>
      </div>
      ${activeClientsSubtab === "list" ? `
        <div class="shows-toolbar-top clients-search-toolbar">
          <label class="sort-control invoice-search-control">
            <span>Search</span>
            <input
              type="search"
              id="clientSearchInput"
              placeholder="Search clients"
              value="${escapeHtml(state.ui.clientSearchQuery || "")}"
            >
          </label>
          <button type="button" class="secondary search-submit-button" id="applyClientSearchButton">Search</button>
        </div>
        <div class="shows-toolbar-top clients-search-toolbar">
          <label class="sort-control">
            <span>GST Details</span>
            <select id="clientGstFilter">
              <option value="all" ${clientGstFilter === "all" ? "selected" : ""}>All Clients</option>
              <option value="missing" ${clientGstFilter === "missing" ? "selected" : ""}>Empty GST Details</option>
              <option value="available" ${clientGstFilter === "available" ? "selected" : ""}>With GST Details</option>
            </select>
          </label>
        </div>
        ${selectedClientDetail ? `
          <section class="detail-card client-detail-panel">
            <header>
              <div>
                <h4>${escapeHtml(getClientDisplayName(selectedClientDetail))}</h4>
                <div class="meta">${selectedClientDetail.contactName || "No contact"}${selectedClientDetail.contactEmail ? ` · ${escapeHtml(selectedClientDetail.contactEmail)}` : ""}${selectedClientDetail.contactPhone ? ` · ${escapeHtml(selectedClientDetail.contactPhone)}` : ""}</div>
                <div class="meta">${selectedClientDetail.state || "State not added yet"}${selectedClientDetail.gstin ? ` · GSTIN: ${escapeHtml(selectedClientDetail.gstin)}` : ""}</div>
              </div>
              <div class="toolbar">
                <button type="button" class="ghost small" data-close-client-detail="true">Close</button>
              </div>
            </header>
            <div class="summary-grid client-detail-summary">
              <div class="summary-card">
                <span class="summary-kicker">Outstanding</span>
                <strong>${formatCurrency(clientOutstanding)}</strong>
                <span class="summary-foot">Open balance across this client’s invoices.</span>
              </div>
              <div class="summary-card">
                <span class="summary-kicker">Collected</span>
                <strong>${formatCurrency(clientPaid)}</strong>
                <span class="summary-foot">Payments received from this client.</span>
              </div>
              <div class="summary-card">
                <span class="summary-kicker">Billed</span>
                <strong>${formatCurrency(clientBilled)}</strong>
                <span class="summary-foot">Total value of all client invoices.</span>
              </div>
              <div class="summary-card">
                <span class="summary-kicker">Relationship</span>
                <strong>${selectedClientShows.length} shows</strong>
                <span class="summary-foot">${selectedClientInvoices.length} invoices · ${selectedClientPayments.length} payments</span>
              </div>
            </div>
            <div class="client-detail-grid">
              <section class="assignment-card">
                <strong>Recent Invoices</strong>
                <div class="stack tight client-detail-list">
                  ${selectedClientInvoices.length ? selectedClientInvoices.slice(0, 5).map((invoice) => `
                    <div class="client-detail-row">
                      <div>
                        <strong>${escapeHtml(invoice.invoiceNumber)}</strong>
                        <div class="meta">${escapeHtml(getInvoiceStatusLabel(invoice))} · ${escapeHtml(formatInvoiceDate(invoice.issueDate))}</div>
                      </div>
                      <strong>${formatCurrency(invoice.totalAmount || 0)}</strong>
                    </div>
                  `).join("") : '<p class="meta">No invoices yet.</p>'}
                </div>
              </section>
              <section class="assignment-card">
                <strong>Payment History</strong>
                <div class="stack tight client-detail-list">
                  ${selectedClientPayments.length ? selectedClientPayments.slice(0, 6).map((payment) => `
                    <div class="client-detail-row">
                      <div>
                        <strong>${escapeHtml(formatInvoiceDate(payment.paymentDate))}</strong>
                        <div class="meta">${escapeHtml(payment.invoiceNumber || "")}${payment.note ? ` · ${escapeHtml(payment.note)}` : ""}</div>
                      </div>
                      <strong>${formatCurrency(payment.amount || 0)}</strong>
                    </div>
                  `).join("") : '<p class="meta">No payments recorded yet.</p>'}
                </div>
              </section>
              <section class="assignment-card">
                <strong>Linked Shows</strong>
                <div class="stack tight client-detail-list">
                  ${selectedClientShows.length ? selectedClientShows.slice(0, 6).map((show) => `
                    <div class="client-detail-row">
                      <div>
                        <strong>${escapeHtml(show.showName)}</strong>
                        <div class="meta">${escapeHtml(formatDateRange(getShowStartDate(show), getShowEndDate(show)))}${show.location ? ` · ${escapeHtml(show.location)}` : ""}</div>
                      </div>
                    </div>
                  `).join("") : '<p class="meta">No linked shows yet.</p>'}
                </div>
              </section>
              <section class="assignment-card">
                <strong>Billing Info</strong>
                <div class="stack tight client-detail-list">
                  <div class="meta">${escapeHtml(selectedClientDetail.billingAddress || "No billing address saved yet.")}</div>
                  ${selectedClientDetail.notes ? `<div class="meta">${escapeHtml(selectedClientDetail.notes)}</div>` : ""}
                </div>
              </section>
            </div>
            <section class="assignment-card client-ledger-section">
              <header>
                <div>
                  <strong>Ledger</strong>
                  <div class="meta">Invoices and payments in running balance order.</div>
                </div>
              </header>
              <div class="client-ledger-table-wrap">
                <table class="client-ledger-table">
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Type</th>
                      <th>Reference</th>
                      <th>Particulars</th>
                      <th>Debit</th>
                      <th>Credit</th>
                      <th>Balance</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${selectedClientLedgerEntries.length ? selectedClientLedgerEntries.map((entry) => `
                      <tr>
                        <td>${escapeHtml(formatInvoiceDate(entry.date))}</td>
                        <td>${escapeHtml(entry.type)}</td>
                        <td>${escapeHtml(entry.reference || "-")}</td>
                        <td>${escapeHtml(entry.particulars || "-")}</td>
                        <td>${entry.debit ? escapeHtml(formatCurrency(entry.debit)) : "-"}</td>
                        <td>${entry.credit ? escapeHtml(formatCurrency(entry.credit)) : "-"}</td>
                        <td>${escapeHtml(formatCurrency(entry.balance))}</td>
                      </tr>
                    `).join("") : `<tr><td colspan="7">No ledger transactions yet.</td></tr>`}
                  </tbody>
                </table>
              </div>
            </section>
          </section>
        ` : ""}
      ` : ""}
      <form id="clientForm" class="stack tight editor-form ${activeClientsSubtab === "create" ? "" : "hidden"}" autocomplete="off">
        <input type="hidden" name="clientId">
        <div class="form-grid editor-section">
          <label class="field"><span>Client Name</span><input type="text" name="clientName" required autocomplete="off" data-form-type="other"></label>
          <label class="field">
            <span>State</span>
            <select name="clientState" data-searchable="true" data-search-placeholder="Search states">
              <option value="">Select state</option>
              ${INDIAN_STATE_OPTIONS.map((stateName) => `<option value="${stateName}">${stateName}</option>`).join("")}
            </select>
          </label>
          <label class="field client-gstin-field">
            <span>GSTIN</span>
            <div class="field-action-row">
              <input type="text" name="clientGstin" autocomplete="off" data-form-type="other">
            </div>
            <div class="field-action-row field-action-row-next">
              <button type="button" class="secondary small" id="fetchGstinDetailsButton">Fetch GST Details</button>
            </div>
          </label>
          <label class="field"><span>Contact Person</span><input type="text" name="contactName" autocomplete="off" data-form-type="other"></label>
          <label class="field"><span>Contact Email</span><input type="email" name="contactEmail" autocomplete="off"></label>
          <label class="field"><span>Contact Phone</span><input type="text" name="contactPhone" autocomplete="off"></label>
        </div>
        <label class="field editor-section"><span>Billing Address</span><textarea name="billingAddress" rows="3"></textarea></label>
        <label class="field editor-section"><span>Notes</span><textarea name="clientNotes" rows="2"></textarea></label>
        <div class="toolbar editor-actions">
          <button type="submit">Save Client</button>
          <button type="button" class="ghost" id="resetClientForm">Clear</button>
        </div>
        <div id="clientFormMessage" class="message"></div>
      </form>
      <div class="stack ${activeClientsSubtab === "list" ? "" : "hidden"}" data-clients-section="list">
        <div class="approval-list">
          ${clientPagination.items.length ? clientPagination.items.map((client) => `
            <article class="show-card">
              <header>
                <div>
                  <h4>${escapeHtml(getClientDisplayName(client))}</h4>
                  <div class="meta">${client.contactName || "No contact"}${client.contactEmail ? ` · ${client.contactEmail}` : ""}${client.contactPhone ? ` · ${client.contactPhone}` : ""}</div>
                  <div class="meta">${client.state || "State not added yet"}</div>
                  <div class="meta">${client.gstin ? `GSTIN: ${client.gstin}` : "GSTIN not added yet"}</div>
                </div>
                <div class="toolbar">
                  <span class="pill">${showsByClientId.get(client.id) || 0} shows</span>
                  <span class="pill">${invoicesByClientId.get(client.id) || 0} invoices</span>
                  <button type="button" class="ghost small" data-view-client="${client.id}">View</button>
                  <button type="button" class="secondary small" data-edit-client="${client.id}">Edit</button>
                  <button type="button" class="ghost small" data-delete-client="${client.id}">Delete</button>
                </div>
              </header>
              <div class="stack tight">
                <div class="meta">${client.billingAddress || "No billing address saved yet."}</div>
                ${client.notes ? `<div class="meta">${client.notes}</div>` : ""}
              </div>
            </article>
          `).join("") : `<p>${clients.length ? "No clients match the current search or GST filter." : "No clients yet. Create the first client above."}</p>`}
        </div>
        ${renderPaginationControls("clients", clientPagination, "clients")}
      </div>
    </div>
  `;

  enhanceCustomSelects(panel);
  wirePaginationControls(panel, "clients", "clientsPage", "clientsPageSize", () => renderClientsPanel());

  const form = document.getElementById("clientForm");
  const message = document.getElementById("clientFormMessage");
  const searchInput = document.getElementById("clientSearchInput");
  const gstFilterSelect = document.getElementById("clientGstFilter");

  const clearClientForm = () => {
    form.reset();
    form.elements.namedItem("clientId").value = "";
    message.textContent = "";
    syncCustomSelect(form.elements.namedItem("clientState"));
  };

  document.getElementById("resetClientForm")?.addEventListener("click", () => {
    if (!confirmDiscardDirtyForm("clear this client form")) return;
    clearDirtyForm("client");
    clearClientForm();
  });

  form.elements.namedItem("clientGstin")?.addEventListener("input", (event) => {
    const stateFromGstin = getStateFromGstin(event.currentTarget.value);
    if (stateFromGstin && !form.elements.namedItem("clientState").value) {
      form.elements.namedItem("clientState").value = stateFromGstin;
      syncCustomSelect(form.elements.namedItem("clientState"));
    }
  });

  document.getElementById("fetchGstinDetailsButton")?.addEventListener("click", async () => {
    const gstin = form.elements.namedItem("clientGstin").value.trim().toUpperCase();
    message.textContent = "";
    if (!gstin) {
      message.textContent = "Enter GSTIN first.";
      return;
    }
    try {
      const result = await apiRequest("/api/admin/gstin-lookup", {
        method: "POST",
        body: JSON.stringify({ gstin })
      });
      form.elements.namedItem("clientGstin").value = result.gstin || gstin;
      if (result.state) {
        form.elements.namedItem("clientState").value = result.state;
        syncCustomSelect(form.elements.namedItem("clientState"));
      }
      if (result.name) {
        form.elements.namedItem("clientName").value = result.name;
      }
      if (result.billingAddress) {
        form.elements.namedItem("billingAddress").value = result.billingAddress;
      }
      if (result.contactEmail) {
        form.elements.namedItem("contactEmail").value = result.contactEmail;
      }
      if (result.contactPhone) {
        form.elements.namedItem("contactPhone").value = result.contactPhone;
      }
      setDirtyForm("client");
      message.textContent = result.message || (result.configured ? "GST details fetched. Please review before saving." : "State filled from GSTIN.");
    } catch (error) {
      const stateFromGstin = getStateFromGstin(gstin);
      if (stateFromGstin) {
        form.elements.namedItem("clientState").value = stateFromGstin;
        syncCustomSelect(form.elements.namedItem("clientState"));
      }
      setDirtyForm("client");
      message.textContent = error.message;
    }
  });

  const applyClientSearch = () => {
    state.ui.clientSearchQuery = searchInput?.value || "";
    state.ui.clientsPage = 1;
    saveState(state);
    renderClientsPanel();
  };
  searchInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    applyClientSearch();
  });
  document.getElementById("applyClientSearchButton")?.addEventListener("click", applyClientSearch);

  gstFilterSelect?.addEventListener("change", (event) => {
    state.ui.clientGstFilter = event.currentTarget.value;
    state.ui.clientsPage = 1;
    saveState(state);
    renderClientsPanel();
  });

  panel.querySelectorAll("[data-clients-subtab]").forEach((button) => {
    button.addEventListener("click", () => {
      if ((button.dataset.clientsSubtab || "list") !== activeClientsSubtab && !confirmDiscardDirtyForm("switch client sections")) {
        return;
      }
      clearDirtyForm();
      state.ui.clientsSubtab = button.dataset.clientsSubtab || "list";
      if (state.ui.clientsSubtab === "create") {
        clearClientForm();
      }
      saveState(state);
      renderClientsPanel();
    });
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const normalized = normalizeClient({
      id: form.elements.namedItem("clientId").value || uid("client"),
      name: form.elements.namedItem("clientName").value,
      state: form.elements.namedItem("clientState").value,
      gstin: form.elements.namedItem("clientGstin").value,
      contactName: form.elements.namedItem("contactName").value,
      contactEmail: form.elements.namedItem("contactEmail").value,
      contactPhone: form.elements.namedItem("contactPhone").value,
      billingAddress: form.elements.namedItem("billingAddress").value,
      notes: form.elements.namedItem("clientNotes").value
    });
    if (!normalized.name) {
      message.textContent = "Client name is required.";
      return;
    }
    const existingIndex = state.clients.findIndex((client) => client.id === normalized.id);
    const isEditingClient = existingIndex >= 0;
    try {
      const payload = await apiRequest("/api/admin/clients", {
        method: "POST",
        body: JSON.stringify(normalized)
      });
      applyServerState(payload);
      clearDirtyForm("client");
      state.ui.clientsSubtab = isEditingClient ? "list" : "create";
      saveState(state);
      renderDashboard();
      if (isEditingClient) {
        showToast("Client updated.");
      } else {
        const nextForm = document.getElementById("clientForm");
        const nextMessage = document.getElementById("clientFormMessage");
        if (nextForm) {
          nextForm.reset();
          nextForm.elements.namedItem("clientId").value = "";
          syncCustomSelect(nextForm.elements.namedItem("clientState"));
        }
        if (nextMessage) {
          nextMessage.textContent = "";
        }
        showToast("Client created. You can add the next one right away.");
      }
    } catch (error) {
      message.textContent = error.message;
      await refreshFromServer();
      renderDashboard();
    }
  });

  panel.addEventListener("click", async (event) => {
    const closeDetailButton = event.target.closest("[data-close-client-detail]");
    if (closeDetailButton) {
      state.ui.selectedClientDetailId = null;
      saveState(state);
      renderClientsPanel();
      return;
    }

    const viewButton = event.target.closest("[data-view-client]");
    if (viewButton) {
      state.ui.selectedClientDetailId = viewButton.dataset.viewClient;
      saveState(state);
      renderClientsPanel();
      return;
    }

    const editButton = event.target.closest("[data-edit-client]");
    if (editButton) {
      if (!confirmDiscardDirtyForm("open another client")) return;
      const client = getClientById(editButton.dataset.editClient);
      if (!client) return;
      clearDirtyForm();
      form.elements.namedItem("clientId").value = client.id;
      form.elements.namedItem("clientName").value = client.name;
      form.elements.namedItem("clientState").value = client.state || "";
      form.elements.namedItem("clientGstin").value = client.gstin || "";
      form.elements.namedItem("contactName").value = client.contactName || "";
      form.elements.namedItem("contactEmail").value = client.contactEmail || "";
      form.elements.namedItem("contactPhone").value = client.contactPhone || "";
      form.elements.namedItem("billingAddress").value = client.billingAddress || "";
      form.elements.namedItem("clientNotes").value = client.notes || "";
      syncCustomSelect(form.elements.namedItem("clientState"));
      state.ui.clientsSubtab = "create";
      saveState(state);
      form.classList.remove("hidden");
      panel.querySelector("[data-clients-section='list']")?.classList.add("hidden");
      panel.querySelectorAll("[data-clients-subtab]").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.clientsSubtab === "create");
      });
      form.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }

    const deleteButton = event.target.closest("[data-delete-client]");
    if (deleteButton) {
      const clientId = deleteButton.dataset.deleteClient;
      const linkedShows = state.shows.some((show) => show.clientId === clientId);
      const linkedInvoices = state.invoices.some((invoice) => invoice.clientId === clientId);
      if (linkedShows || linkedInvoices) {
        const showCount = state.shows.filter((show) => show.clientId === clientId).length;
        const invoiceCount = state.invoices.filter((invoice) => invoice.clientId === clientId).length;
        const shouldDelete = window.confirm(
          `This client is linked to ${showCount} ${showCount === 1 ? "show" : "shows"} and ${invoiceCount} ${invoiceCount === 1 ? "invoice" : "invoices"}.\n\nDelete it from the client master anyway? Existing shows and invoices will keep the client name for history, but they will no longer be linked to this client record.`
        );
        if (!shouldDelete) return;
      } else {
        const client = state.clients.find((item) => item.id === clientId);
        const shouldDelete = window.confirm(`Delete client "${client?.name || "this client"}"?`);
        if (!shouldDelete) return;
      }
      try {
        const payload = await apiRequest(`/api/admin/clients/${encodeURIComponent(clientId)}${linkedShows || linkedInvoices ? "?keepHistory=true" : ""}`, {
          method: "DELETE"
        });
        applyServerState(payload);
        saveState(state);
        renderDashboard();
        showToast("Client deleted.");
      } catch (error) {
        showToast(error.message);
        await refreshFromServer();
        renderDashboard();
      }
    }
  });
  wireDirtyFormTracking(form, "client");
}

function renderShowForm() {
  const panel = document.getElementById("showFormPanel");
  const crewOptions = getCrewUsers();
  const adminOptions = getApprovedAdminUsers();
  const clientOptions = getSortedClients();
  ensureUiState();
  const editingShow = state.ui.editingShowId
    ? state.shows.find((show) => show.id === state.ui.editingShowId)
    : null;
  const draftShow = !editingShow && state.ui.showDraftTemplate && typeof state.ui.showDraftTemplate === "object"
    ? state.ui.showDraftTemplate
    : null;
  const isEditing = Boolean(editingShow);
  const defaultShowDate = !isEditing && state.ui.newShowDate ? state.ui.newShowDate : "";
  const showDateFromValue = isEditing ? getShowStartDate(editingShow) : (draftShow?.showDateFrom || defaultShowDate);
  const showDateToValue = isEditing ? getShowEndDate(editingShow) : (draftShow?.showDateTo || defaultShowDate);
  const selectedClientId = isEditing
    ? (editingShow.clientId || getClientByName(editingShow.client)?.id || "")
    : (draftShow?.clientId || "");
  const selectedShowStatus = isEditing
    ? (editingShow.showStatus === "tentative" ? "tentative" : "confirmed")
    : (draftShow?.showStatus === "tentative" ? "tentative" : "confirmed");
  const selectedLightDesignerId = isEditing
    ? (editingShow.assignments || []).find((assignment) => assignment.lightDesignerId)?.lightDesignerId || ""
    : (draftShow?.showLightDesignerId || "");

  panel.innerHTML = `
    <div class="stack">
      <div>
        <div class="form-header">
          <div>
            <h3>${isEditing ? `Editing: ${editingShow.showName}` : draftShow ? `Duplicate Draft: ${draftShow.showName}` : "Create New Show"}</h3>
            <p class="muted-note">Only admins can edit show amount and operator amounts.</p>
          </div>
          ${isEditing ? '<span class="pill edit-pill">Edit Mode</span>' : ""}
        </div>
      </div>
      <form id="showForm" class="stack tight editor-form" autocomplete="off">
        <input type="hidden" name="showId" value="${escapeHtml(isEditing ? editingShow.id : "")}">
        <div class="form-grid editor-section">
          <label class="field"><span>Show Date From</span><input type="date" name="showDateFrom" value="${escapeHtml(showDateFromValue)}" required autocomplete="off"></label>
          <label class="field"><span>Show Date To</span><input type="date" name="showDateTo" value="${escapeHtml(showDateToValue)}" required autocomplete="off"></label>
          <label class="field"><span>Show Name</span><input type="text" name="showName" value="${escapeHtml(isEditing ? editingShow.showName : (draftShow?.showName || ""))}" required autocomplete="off" data-form-type="other"></label>
          <label class="field">
            <span>Client</span>
            <select name="clientId" ${clientOptions.length ? "required" : ""} data-searchable="true" data-search-placeholder="Search clients">
              <option value="">${clientOptions.length ? "Select client" : "No clients yet"}</option>
              ${clientOptions.map((client) => `<option value="${client.id}" ${selectedClientId === client.id ? "selected" : ""}>${escapeHtml(getClientDisplayName(client))}</option>`).join("")}
            </select>
          </label>
          <label class="field">
            <span>Light Designer</span>
            <select name="showLightDesignerId">
              <option value="">Select admin</option>
              ${adminOptions.map((adminUser) => `<option value="${adminUser.id}" ${selectedLightDesignerId === adminUser.id ? "selected" : ""}>${adminUser.name}</option>`).join("")}
            </select>
          </label>
          <label class="field"><span>Location</span><input type="text" name="location" value="${escapeHtml(isEditing ? editingShow.location : (draftShow?.location || ""))}" autocomplete="off" data-form-type="other"></label>
          <label class="field"><span>Amount of the Show</span><input type="number" name="amountShow" value="${escapeHtml(isEditing ? editingShow.amountShow : (draftShow?.amountShow ?? ""))}" min="0" step="1" autocomplete="off"></label>
        </div>
        ${clientOptions.length ? "" : '<p class="muted-note">Add clients in the Clients tab before creating shows.</p>'}
        <div class="field editor-section">
          <span>Show Status</span>
          <div class="status-toggle" role="radiogroup" aria-label="Show Status">
            <label class="status-option">
              <input type="radio" name="showStatus" value="confirmed" ${selectedShowStatus === "confirmed" ? "checked" : ""}>
              <span>Confirmed</span>
            </label>
            <label class="status-option">
              <input type="radio" name="showStatus" value="tentative" ${selectedShowStatus === "tentative" ? "checked" : ""}>
              <span>Tentative</span>
            </label>
          </div>
        </div>
        <div class="stack tight editor-section">
          <strong>Assign Crew Members</strong>
          <div id="assignmentEditor" class="stack tight"></div>
          <button type="button" class="secondary" id="addAssignmentRow">Add Crew Assignment</button>
        </div>
        <div class="toolbar editor-actions">
          <button type="submit">${isEditing ? "Update Show" : "Save Show"}</button>
          <button type="button" class="ghost" id="resetShowForm">${isEditing ? "Cancel Edit" : "Clear"}</button>
          ${isEditing ? '<button type="button" class="danger" id="deleteShowButton">Delete Show</button>' : ""}
        </div>
      </form>
    </div>
  `;

  const assignmentEditor = document.getElementById("assignmentEditor");
  enhanceCustomSelects(panel);

  function addAssignmentRow(assignment = null) {
    const row = document.createElement("div");
    row.className = "form-grid";
    row.dataset.assignmentRow = "true";
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
        <input type="number" name="assignmentAmount" min="0" step="1" value="${assignment?.operatorAmount ?? ""}" autocomplete="off">
      </label>
      <label class="field">
        <span>Onward Travel Date</span>
        <input type="date" name="assignmentOnwardTravelDate" value="${assignment?.onwardTravelDate ?? ""}" autocomplete="off">
      </label>
      <label class="field">
        <span>Return Travel Date</span>
        <input type="date" name="assignmentReturnTravelDate" value="${assignment?.returnTravelDate ?? ""}" autocomplete="off">
      </label>
      <label class="field">
        <span>Onward Travel Sector</span>
        <input type="text" name="assignmentOnwardTravelSector" value="${assignment?.onwardTravelSector ?? ""}" autocomplete="off" data-form-type="other">
      </label>
      <label class="field">
        <span>Return Travel Sector</span>
        <input type="text" name="assignmentReturnTravelSector" value="${assignment?.returnTravelSector ?? ""}" autocomplete="off" data-form-type="other">
      </label>
      <label class="field">
        <span>Notes</span>
        <input type="text" name="assignmentNotes" value="${assignment?.notes ?? ""}" autocomplete="off" data-form-type="other">
      </label>
      <button type="button" class="ghost small remove-assignment">Remove</button>
    `;
    assignmentEditor.append(row);
    enhanceCustomSelects(row);
    row.querySelector(".remove-assignment").addEventListener("click", () => {
      const confirmed = window.confirm("Remove this crew assignment?");
      if (!confirmed) return;
      row.remove();
      setDirtyForm("show");
    });
  }

  const initialAssignments = isEditing
    ? (editingShow.assignments || [])
    : (draftShow?.assignments || []);

  if (initialAssignments.length) {
    initialAssignments.forEach((assignment) => addAssignmentRow(assignment));
  } else {
    addAssignmentRow();
  }

  document.getElementById("addAssignmentRow").addEventListener("click", () => {
    addAssignmentRow();
    setDirtyForm("show");
  });
  document.getElementById("resetShowForm").addEventListener("click", () => {
    if (!confirmDiscardDirtyForm("leave this show form")) return;
    clearDirtyForm("show");
    const returnTab = getShowReturnTab();
    resetEditingState();
    state.ui.activeSidebarTab = returnTab;
    state.ui.showSubtab = returnTab === "showsPanel" ? "create" : "list";
    saveState(state);
    renderDashboard();
    restoreShowReturnContext();
  });

  if (isEditing) {
    document.getElementById("deleteShowButton").addEventListener("click", async () => {
      const confirmDelete = window.confirm(`Delete "${editingShow.showName}"?`);
      if (!confirmDelete) return;
      state.shows = state.shows.filter((show) => show.id !== editingShow.id);
      const returnTab = getShowReturnTab();
      clearDirtyForm("show");
      resetEditingState();
      try {
        await syncAdminState();
        state.ui.activeSidebarTab = returnTab;
        state.ui.showSubtab = returnTab === "showsPanel" ? "list" : "list";
        saveState(state);
        renderDashboard();
        restoreShowReturnContext();
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
    const showLightDesignerId = formData.get("showLightDesignerId").toString();
    const assignmentRows = [...assignmentEditor.querySelectorAll('[data-assignment-row="true"]')];
    const assignments = assignmentRows.map((row) => {
      const crewSelect = row.querySelector('select[name="assignmentCrew"]');
      const operatorAmountInput = row.querySelector('input[name="assignmentAmount"]');
      const onwardTravelDateInput = row.querySelector('input[name="assignmentOnwardTravelDate"]');
      const returnTravelDateInput = row.querySelector('input[name="assignmentReturnTravelDate"]');
      const onwardTravelSectorInput = row.querySelector('input[name="assignmentOnwardTravelSector"]');
      const returnTravelSectorInput = row.querySelector('input[name="assignmentReturnTravelSector"]');
      const notesInput = row.querySelector('input[name="assignmentNotes"]');
      const crewValue = crewSelect?.value || "";

      return {
        crewId: crewValue,
        manualCrewName: "",
        lightDesignerId: showLightDesignerId,
        operatorAmount: operatorAmountInput?.value || "",
        onwardTravelDate: onwardTravelDateInput?.value || "",
        returnTravelDate: returnTravelDateInput?.value || "",
        onwardTravelSector: onwardTravelSectorInput?.value || "",
        returnTravelSector: returnTravelSectorInput?.value || "",
        notes: notesInput?.value || ""
      };
    })
      .filter((assignment) => assignment.crewId || assignment.manualCrewName)
      .map((assignment) => ({
        crewId: assignment.crewId,
        manualCrewName: assignment.manualCrewName,
        lightDesignerId: assignment.lightDesignerId,
        operatorAmount: Number(assignment.operatorAmount || 0),
        onwardTravelDate: assignment.onwardTravelDate,
        returnTravelDate: assignment.returnTravelDate,
        onwardTravelSector: assignment.onwardTravelSector.trim(),
        returnTravelSector: assignment.returnTravelSector.trim(),
        notes: assignment.notes.trim()
      }));

    const uniqueCrewKeys = new Set(assignments.map((assignment) => assignment.crewId || `manual:${assignment.manualCrewName.toLowerCase()}`));
    if (uniqueCrewKeys.size !== assignments.length) {
      alert("Each crew member can only be assigned once per show.");
      return;
    }

    const payload = {
      id: showId || uid("show"),
      showDateFrom: formData.get("showDateFrom").toString(),
      showDateTo: formData.get("showDateTo").toString(),
      showDate: formData.get("showDateFrom").toString(),
      showStatus: formData.get("showStatus").toString() === "tentative" ? "tentative" : "confirmed",
      googleEventId: editingShow?.googleEventId || "",
      googleSyncSource: "pixelbug",
      googleSyncStatus: "pending_push",
      googleNotes: editingShow?.googleNotes || "",
      googleLastSyncedAt: editingShow?.googleLastSyncedAt || "",
      needsAdminCompletion: editingShow?.googleEventId ? false : Boolean(editingShow?.needsAdminCompletion),
      showName: formData.get("showName").toString().trim(),
      clientId: formData.get("clientId").toString().trim(),
      client: "",
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

    const selectedClient = getClientById(payload.clientId);
    if (!selectedClient) {
      alert("Please select a client from the Clients tab.");
      return;
    }
    payload.client = selectedClient.name;

    const existingIndex = state.shows.findIndex((show) => show.id === payload.id);
    if (existingIndex >= 0) {
      state.shows[existingIndex] = payload;
    } else {
      state.shows.push(payload);
    }

    clearDirtyForm("show");
    resetEditingState();
    try {
      await syncAdminState();
      const returnTab = getShowReturnTab();
      state.ui.activeSidebarTab = returnTab;
      state.ui.showSubtab = returnTab === "showsPanel" ? "list" : "list";
      saveState(state);
      renderDashboard();
      restoreShowReturnContext();
      showToast(existingIndex >= 0 ? "Show updated." : "Show created.");
    } catch (error) {
      showToast(error.message);
      await refreshFromServer();
      renderDashboard();
    }
  });
  wireDirtyFormTracking(document.getElementById("showForm"), "show");
}

function fillShowForm(showId, returnTab = state.ui.activeSidebarTab || "showsPanel") {
  if (!confirmDiscardDirtyForm("open another show")) return;
  const show = state.shows.find((item) => item.id === showId);
  if (!show) return;
  ensureUiState();
  clearDirtyForm();
  captureShowReturnContext(returnTab);
  state.ui.editingShowId = showId;
  state.ui.newShowDate = "";
  state.ui.showDraftTemplate = null;
  state.ui.showSubtab = "create";
  state.ui.activeSidebarTab = "showsPanel";
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
  setFieldValue("clientId", show.clientId || getClientByName(show.client)?.id || "");
  setFieldValue("showLightDesignerId", (show.assignments || []).find((assignment) => assignment.lightDesignerId)?.lightDesignerId || "");
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
    row.querySelector('select[name="assignmentCrew"]').value = assignment.crewId || "";
    row.querySelector('select[name="assignmentCrew"]').dispatchEvent(new Event("change", { bubbles: true }));
    row.querySelector('input[name="assignmentAmount"]').value = assignment.operatorAmount;
    row.querySelector('input[name="assignmentOnwardTravelDate"]').value = assignment.onwardTravelDate || "";
    row.querySelector('input[name="assignmentReturnTravelDate"]').value = assignment.returnTravelDate || "";
    row.querySelector('input[name="assignmentOnwardTravelSector"]').value = assignment.onwardTravelSector || "";
    row.querySelector('input[name="assignmentReturnTravelSector"]').value = assignment.returnTravelSector || "";
    row.querySelector('input[name="assignmentNotes"]').value = assignment.notes || "";
  });

  form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function startShowDraftFromExistingShow(showId, returnTab = state.ui.activeSidebarTab || "showsPanel") {
  if (!confirmDiscardDirtyForm("duplicate this show")) return;
  const show = state.shows.find((item) => item.id === showId);
  if (!show) return;
  ensureUiState();
  clearDirtyForm();
  captureShowReturnContext(returnTab);
  state.ui.editingShowId = null;
  state.ui.newShowDate = "";
  state.ui.showDraftTemplate = {
    showDateFrom: getShowStartDate(show),
    showDateTo: getShowEndDate(show),
    showName: show.showName,
    clientId: show.clientId || getClientByName(show.client)?.id || "",
    location: show.location || "",
    amountShow: show.amountShow || 0,
    showStatus: show.showStatus === "tentative" ? "tentative" : "confirmed",
    showLightDesignerId: (show.assignments || []).find((assignment) => assignment.lightDesignerId)?.lightDesignerId || "",
    assignments: (show.assignments || []).map((assignment) => ({
      crewId: assignment.crewId || "",
      lightDesignerId: assignment.lightDesignerId || "",
      operatorAmount: assignment.operatorAmount ?? "",
      onwardTravelDate: assignment.onwardTravelDate || "",
      returnTravelDate: assignment.returnTravelDate || "",
      onwardTravelSector: assignment.onwardTravelSector || "",
      returnTravelSector: assignment.returnTravelSector || "",
      notes: assignment.notes || ""
    }))
  };
  state.ui.showSubtab = "create";
  state.ui.activeSidebarTab = "showsPanel";
  saveState(state);
  renderSidebarTabs();
  renderDashboard();
  showToast("Show duplicated into a new draft.");
}

function startShowDraftForDate(showDate) {
  ensureUiState();
  captureShowReturnContext("calendarPanel");
  state.ui.editingShowId = null;
  state.ui.newShowDate = showDate || dateKey(new Date());
  state.ui.showDraftTemplate = null;
  state.ui.showSubtab = "create";
  state.ui.activeSidebarTab = "showsPanel";
  saveState(state);
  renderSidebarTabs();
  renderDashboard();
  showToast("Create Show opened for selected date.");
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
              <span class="pill">${getRoleLabel(user.role)}</span>
            </header>
            ${(user.role === "crew" || user.role === "admin") && user.color ? `<div class="meta">Requested color <span class="legend-swatch" style="background:${resolveCrewColor(user.color)}"></span>${resolveCrewColor(user.color)}</div>` : ""}
            <div class="toolbar" style="margin-top:12px;">
              <button type="button" class="small" data-approve="${user.id}">Approve</button>
              <button type="button" class="secondary small" data-reject="${user.id}">Reject</button>
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
  const approvedAccounts = sortUsersByName(getApprovedAccountsUsers());
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
      <form id="adminAccountsCreateForm" class="stack tight">
        <div class="form-header">
          <div>
            <h3>Add Accounts Login</h3>
            <p class="muted-note">Accounts users can log in separately and access the Invoices tab without full admin controls.</p>
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
          <button type="submit">Add Accounts User</button>
        </div>
        <div id="adminAccountsMessage" class="message"></div>
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
                <span class="legend-swatch" style="background:${resolveCrewColor(member.color) || "#264653"}"></span>
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
                <span class="pill" style="background:${resolveCrewColor(member.color)}; color:white;">Crew</span>
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
            <h3>Approved Accounts Logins</h3>
            <p class="muted-note">Accounts users can manage invoicing and collections without crew or Google admin powers.</p>
          </div>
          <span class="pill">${approvedAccounts.length} accounts</span>
        </div>
        <div class="approval-list">
          ${approvedAccounts.length ? approvedAccounts.map((member) => `
            <article class="show-card">
              <header>
                <div>
                  <strong>${member.name}</strong>
                  <div class="meta">${member.email} · ${member.phone}</div>
                </div>
                <span class="pill">Accounts</span>
              </header>
              <div class="toolbar" style="margin-top:12px;">
                <button type="button" class="danger small" data-remove-accounts="${member.id}">Remove Accounts Login</button>
              </div>
            </article>
          `).join("") : "<p>No approved accounts logins yet.</p>"}
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
  const accountsForm = document.getElementById("adminAccountsCreateForm");
  const accountsMessage = document.getElementById("adminAccountsMessage");
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

  if (accountsForm) {
    accountsForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const data = new FormData(accountsForm);
      const email = data.get("email").toString().trim().toLowerCase();

      if (state.users.some((user) => user.email.toLowerCase() === email)) {
        accountsMessage.textContent = "That email already exists.";
        return;
      }

      try {
        const payload = await apiRequest("/api/admin/add-accounts", {
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
        showToast("Accounts login added.");
      } catch (error) {
        accountsMessage.textContent = error.message;
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

  document.querySelectorAll("[data-remove-accounts]").forEach((button) => {
    button.addEventListener("click", async () => {
      const accountsId = button.dataset.removeAccounts;
      const accountsUser = getUserById(accountsId);
      if (!accountsId || !accountsUser) return;
      const confirmed = window.confirm(`Remove accounts login "${accountsUser.name}"?`);
      if (!confirmed) return;

      state.users = state.users.filter((user) => user.id !== accountsId);
      try {
        await syncAdminState();
        saveState(state);
        renderDashboard();
        showToast("Accounts login removed.");
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
  applyThemeFromState();
  attachThemeListener();
  setupGoogleAutoRefresh();
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
  window.addEventListener("beforeunload", (event) => {
    if (!hasDirtyForm()) return;
    event.preventDefault();
    event.returnValue = "";
  });
  try {
    await refreshFromServer();
  } catch (error) {
    showToast("Server connection failed.");
  }
  const url = new URL(window.location.href);
  const googleStatus = url.searchParams.get("google");
  if (googleStatus === "connected") {
    showToast("Google Calendar connected.");
    url.searchParams.delete("google");
    window.history.replaceState({}, document.title, `${url.pathname}${url.search}`);
  } else if (googleStatus === "error") {
    showToast("Google Calendar connection failed.");
    url.searchParams.delete("google");
    window.history.replaceState({}, document.title, `${url.pathname}${url.search}`);
  }
  render();
}

initApp();
