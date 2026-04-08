const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { DatabaseSync } = require("node:sqlite");

function loadLocalEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) return;
    const equalsIndex = trimmed.indexOf("=");
    if (equalsIndex === -1) return;
    const key = trimmed.slice(0, equalsIndex).trim();
    const value = trimmed.slice(equalsIndex + 1).trim().replace(/^['"]|['"]$/g, "");
    if (key && !process.env[key]) {
      process.env[key] = value;
    }
  });
}

loadLocalEnvFile(path.join(__dirname, ".env.local"));
loadLocalEnvFile(path.join(__dirname, ".env"));

const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const DATA_DIR = process.env.PIXELBUG_DATA_DIR
  ? path.resolve(process.env.PIXELBUG_DATA_DIR)
  : path.join(ROOT, "data");
const DATA_FILE = process.env.PIXELBUG_STORE_PATH
  ? path.resolve(process.env.PIXELBUG_STORE_PATH)
  : path.join(DATA_DIR, "store.json");
const DB_FILE = process.env.PIXELBUG_DB_PATH
  ? path.resolve(process.env.PIXELBUG_DB_PATH)
  : path.join(DATA_DIR, "pixelbug.db");
const OUTBOX_FILE = process.env.PIXELBUG_OUTBOX_PATH
  ? path.resolve(process.env.PIXELBUG_OUTBOX_PATH)
  : path.join(DATA_DIR, "email-outbox.log");
const SESSION_COOKIE = "pixelbug_session";
const GOOGLE_TOKEN_SETTING_KEY = "google_oauth_tokens";
const GOOGLE_LAST_SYNC_SETTING_KEY = "google_last_sync_at";
const GOOGLE_LAST_ERROR_SETTING_KEY = "google_last_sync_error";
const GOOGLE_SYNC_DEBOUNCE_MS = 1000 * 30;
const INVOICE_STATUSES = new Set(["draft", "sent", "partially_paid", "paid", "cancelled"]);
const SELF_SERVICE_ROLES = new Set(["admin", "crew", "viewer"]);
const sessions = new Map();
const googleOauthStates = new Map();
let db;
const IS_PRODUCTION = process.env.NODE_ENV === "production";
const COOKIE_SECURE = process.env.PIXELBUG_COOKIE_SECURE === "true" || IS_PRODUCTION;
const GST_LOOKUP_API_URL = process.env.GST_LOOKUP_API_URL || "";
const GST_LOOKUP_API_KEY = process.env.GST_LOOKUP_API_KEY || "";
const GST_LOOKUP_API_KEY_HEADER = process.env.GST_LOOKUP_API_KEY_HEADER || "x-api-key";
const CLEARTAX_GST_LOOKUP_URL = process.env.CLEARTAX_GST_LOOKUP_URL || "https://cleartax.in/f/compliance-report";
const CLEARTAX_GST_LOOKUP_ENABLED = process.env.CLEARTAX_GST_LOOKUP_ENABLED !== "false";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

const GST_STATE_OPTIONS = [
  "Jammu and Kashmir (01)",
  "Himachal Pradesh (02)",
  "Punjab (03)",
  "Chandigarh (04)",
  "Uttarakhand (05)",
  "Haryana (06)",
  "Delhi (07)",
  "Rajasthan (08)",
  "Uttar Pradesh (09)",
  "Bihar (10)",
  "Sikkim (11)",
  "Arunachal Pradesh (12)",
  "Nagaland (13)",
  "Manipur (14)",
  "Mizoram (15)",
  "Tripura (16)",
  "Meghalaya (17)",
  "Assam (18)",
  "West Bengal (19)",
  "Jharkhand (20)",
  "Odisha (21)",
  "Chhattisgarh (22)",
  "Madhya Pradesh (23)",
  "Gujarat (24)",
  "Dadra and Nagar Haveli and Daman and Diu (26)",
  "Maharashtra (27)",
  "Karnataka (29)",
  "Goa (30)",
  "Lakshadweep (31)",
  "Kerala (32)",
  "Tamil Nadu (33)",
  "Puducherry (34)",
  "Andaman and Nicobar Islands (35)",
  "Telangana (36)",
  "Andhra Pradesh (37)",
  "Ladakh (38)"
];

function uid(prefix) {
  return `${prefix}_${crypto.randomBytes(6).toString("hex")}`;
}

function getStateFromGstin(gstin) {
  const code = String(gstin || "").trim().slice(0, 2);
  if (!/^\d{2}$/.test(code)) return "";
  return GST_STATE_OPTIONS.find((stateName) => stateName.endsWith(`(${code})`)) || "";
}

function getNestedValue(source, paths) {
  for (const pathKey of paths) {
    const value = pathKey.split(".").reduce((current, key) => current && current[key], source);
    if (value !== undefined && value !== null && String(value).trim()) {
      return value;
    }
  }
  return "";
}

function normalizeGstinLookupPayload(gstin, payload = {}) {
  const source = payload?.data && typeof payload.data === "object"
    ? payload.data
    : payload?.taxpayerInfo && typeof payload.taxpayerInfo === "object"
      ? payload.taxpayerInfo
      : payload;
  const addressParts = [
    getNestedValue(source, ["pradr.addr.bno", "principalPlaceOfBusiness.address.buildingNumber"]),
    getNestedValue(source, ["pradr.addr.flno", "principalPlaceOfBusiness.address.floorNumber"]),
    getNestedValue(source, ["pradr.addr.bnm", "principalPlaceOfBusiness.address.buildingName"]),
    getNestedValue(source, ["pradr.addr.st", "principalPlaceOfBusiness.address.street"]),
    getNestedValue(source, ["pradr.addr.loc", "principalPlaceOfBusiness.address.location"]),
    getNestedValue(source, ["pradr.addr.dst", "principalPlaceOfBusiness.address.district"]),
    getNestedValue(source, ["pradr.addr.stcd", "principalPlaceOfBusiness.address.state"]),
    getNestedValue(source, ["pradr.addr.pncd", "principalPlaceOfBusiness.address.pincode"])
  ].filter(Boolean);

  return {
    gstin,
    state: getStateFromGstin(gstin),
    name: String(getNestedValue(source, ["lgnm", "legalName", "legal_name", "taxpayer.legalName"]) || getNestedValue(source, ["tradeNam", "tradeName", "trade_name"]) || "").trim(),
    billingAddress: String(getNestedValue(source, ["address", "billingAddress", "principalAddress"]) || addressParts.join(", ")).trim(),
    contactEmail: String(getNestedValue(source, ["email", "contactEmail", "contact.email"]) || "").trim(),
    contactPhone: String(getNestedValue(source, ["mobile", "phone", "contactPhone", "contact.mobile"]) || "").trim(),
    tradeName: String(getNestedValue(source, ["tradeNam", "tradeName", "trade_name"]) || "").trim(),
    status: String(getNestedValue(source, ["sts", "status", "gstinStatus"]) || "").trim()
  };
}

function hasGstinLookupDetails(result) {
  return Boolean(result?.name || result?.billingAddress || result?.tradeName || result?.status);
}

async function lookupConfiguredGstinApi(normalizedGstin) {
  if (!GST_LOOKUP_API_URL) return null;
  const headers = { "Content-Type": "application/json" };
  if (GST_LOOKUP_API_KEY) {
    headers[GST_LOOKUP_API_KEY_HEADER] = GST_LOOKUP_API_KEY;
  }
  const response = await fetch(GST_LOOKUP_API_URL, {
    method: "POST",
    headers,
    body: JSON.stringify({ gstin: normalizedGstin })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || payload.message || "GST lookup failed.");
  }
  return {
    ...normalizeGstinLookupPayload(normalizedGstin, payload),
    configured: true,
    source: "configured-api",
    message: "GST details fetched. Please review before saving."
  };
}

async function lookupCleartaxGstinDetails(normalizedGstin) {
  if (!CLEARTAX_GST_LOOKUP_ENABLED) return null;
  const endpoint = `${CLEARTAX_GST_LOOKUP_URL.replace(/\/$/, "")}/${encodeURIComponent(normalizedGstin)}/?captcha_token=`;
  const response = await fetch(endpoint, {
    headers: {
      Accept: "application/json,text/plain,*/*",
      "User-Agent": "PixelBug GST lookup (https://www.pixelbug.in)"
    }
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) return null;
  const normalized = normalizeGstinLookupPayload(normalizedGstin, payload);
  if (!hasGstinLookupDetails(normalized)) return null;
  return {
    ...normalized,
    configured: true,
    source: "cleartax",
    message: "GST details fetched from ClearTax. Please review before saving."
  };
}

async function lookupGstinDetails(gstin) {
  const normalizedGstin = String(gstin || "").trim().toUpperCase();
  if (!/^\d{2}[A-Z0-9]{13}$/.test(normalizedGstin)) {
    throw new Error("Enter a valid 15-character GSTIN.");
  }

  const state = getStateFromGstin(normalizedGstin);

  const configuredResult = await lookupConfiguredGstinApi(normalizedGstin);
  if (hasGstinLookupDetails(configuredResult)) {
    return configuredResult;
  }

  const cleartaxResult = await lookupCleartaxGstinDetails(normalizedGstin).catch(() => null);
  if (hasGstinLookupDetails(cleartaxResult)) {
    return cleartaxResult;
  }

  return {
    gstin: normalizedGstin,
    state,
    configured: false,
    message: "State filled from GSTIN. ClearTax did not return public details for this GSTIN, so please fill the remaining client info manually."
  };
}

function ensureDataStore() {
  const requiredDirs = new Set([DATA_DIR, path.dirname(DATA_FILE), path.dirname(DB_FILE), path.dirname(OUTBOX_FILE)]);
  requiredDirs.forEach((dirPath) => {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
  });

  if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify({ users: [], shows: [], clients: [] }, null, 2));
  }
}

function ensureDatabase() {
  ensureDataStore();
  if (db) return db;

  db = new DatabaseSync(DB_FILE);
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE,
      phone TEXT NOT NULL,
      role TEXT NOT NULL,
      approved INTEGER NOT NULL DEFAULT 0,
      color TEXT,
      email_verified INTEGER NOT NULL DEFAULT 1,
      verification_token TEXT,
      reset_token TEXT,
      reset_token_expires_at INTEGER,
      password_hash TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS shows (
      id TEXT PRIMARY KEY,
      show_date TEXT NOT NULL,
      show_date_from TEXT,
      show_date_to TEXT,
      show_status TEXT NOT NULL DEFAULT 'confirmed',
      google_event_id TEXT,
      google_sync_source TEXT,
      google_sync_status TEXT,
      google_notes TEXT,
      google_last_synced_at TEXT,
      needs_admin_completion INTEGER NOT NULL DEFAULT 0,
      google_archived INTEGER NOT NULL DEFAULT 0,
      google_archived_at TEXT,
      google_pinned INTEGER NOT NULL DEFAULT 0,
      show_name TEXT NOT NULL,
      client_id TEXT,
      client TEXT,
      venue TEXT,
      location TEXT,
      show_time TEXT,
      amount_show REAL NOT NULL DEFAULT 0,
      assignments_json TEXT NOT NULL DEFAULT '[]'
    );
    CREATE TABLE IF NOT EXISTS clients (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      state TEXT,
      billing_address TEXT,
      gstin TEXT,
      contact_name TEXT,
      contact_email TEXT,
      contact_phone TEXT,
      notes TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
    CREATE TABLE IF NOT EXISTS invoices (
      id TEXT PRIMARY KEY,
      invoice_number TEXT NOT NULL UNIQUE,
      client_name TEXT NOT NULL,
      client_id TEXT,
      issue_date TEXT NOT NULL,
      due_date TEXT,
      status TEXT NOT NULL DEFAULT 'draft',
      notes TEXT,
      details_json TEXT NOT NULL DEFAULT '{}',
      tax_percent REAL NOT NULL DEFAULT 0,
      subtotal REAL NOT NULL DEFAULT 0,
      tax_amount REAL NOT NULL DEFAULT 0,
      total_amount REAL NOT NULL DEFAULT 0,
      amount_paid REAL NOT NULL DEFAULT 0,
      currency TEXT NOT NULL DEFAULT 'INR',
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS invoice_line_items (
      id TEXT PRIMARY KEY,
      invoice_id TEXT NOT NULL,
      show_id TEXT,
      description TEXT NOT NULL,
      sac TEXT,
      custom_details TEXT,
      discount TEXT,
      discount_amount REAL NOT NULL DEFAULT 0,
      quantity REAL NOT NULL DEFAULT 1,
      unit_rate REAL NOT NULL DEFAULT 0,
      amount REAL NOT NULL DEFAULT 0,
      line_order INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS invoice_payments (
      id TEXT PRIMARY KEY,
      invoice_id TEXT NOT NULL,
      payment_date TEXT NOT NULL,
      amount REAL NOT NULL DEFAULT 0,
      note TEXT,
      created_at TEXT NOT NULL,
      FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
    );
  `);

  const existingTables = new Set(
    db.prepare("SELECT name FROM sqlite_master WHERE type = 'table'").all().map((row) => row.name)
  );
  if (!existingTables.has("clients")) {
    db.exec(`
      CREATE TABLE clients (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        billing_address TEXT,
        gstin TEXT,
        contact_name TEXT,
        contact_email TEXT,
        contact_phone TEXT,
        notes TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      )
    `);
  }
  if (!existingTables.has("app_settings")) {
    db.exec(`
      CREATE TABLE app_settings (
        key TEXT PRIMARY KEY,
        value TEXT
      )
    `);
  }
  if (!existingTables.has("invoices")) {
    db.exec(`
      CREATE TABLE invoices (
        id TEXT PRIMARY KEY,
        invoice_number TEXT NOT NULL UNIQUE,
        client_name TEXT NOT NULL,
        client_id TEXT,
        issue_date TEXT NOT NULL,
        due_date TEXT,
        status TEXT NOT NULL DEFAULT 'draft',
        notes TEXT,
        details_json TEXT NOT NULL DEFAULT '{}',
        tax_percent REAL NOT NULL DEFAULT 0,
        subtotal REAL NOT NULL DEFAULT 0,
        tax_amount REAL NOT NULL DEFAULT 0,
        total_amount REAL NOT NULL DEFAULT 0,
        amount_paid REAL NOT NULL DEFAULT 0,
        currency TEXT NOT NULL DEFAULT 'INR',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      )
    `);
  }
  if (!existingTables.has("invoice_line_items")) {
    db.exec(`
      CREATE TABLE invoice_line_items (
        id TEXT PRIMARY KEY,
        invoice_id TEXT NOT NULL,
        show_id TEXT,
        description TEXT NOT NULL,
        sac TEXT,
        custom_details TEXT,
        discount TEXT,
        discount_amount REAL NOT NULL DEFAULT 0,
        quantity REAL NOT NULL DEFAULT 1,
        unit_rate REAL NOT NULL DEFAULT 0,
        amount REAL NOT NULL DEFAULT 0,
        line_order INTEGER NOT NULL DEFAULT 0,
        FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
      )
    `);
  }
  if (!existingTables.has("invoice_payments")) {
    db.exec(`
      CREATE TABLE invoice_payments (
        id TEXT PRIMARY KEY,
        invoice_id TEXT NOT NULL,
        payment_date TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        note TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
      )
    `);
  }

  const showColumns = db.prepare("PRAGMA table_info(shows)").all().map((column) => column.name);
  if (!showColumns.includes("show_date_from")) {
    db.exec("ALTER TABLE shows ADD COLUMN show_date_from TEXT");
  }
  if (!showColumns.includes("show_date_to")) {
    db.exec("ALTER TABLE shows ADD COLUMN show_date_to TEXT");
  }
  if (!showColumns.includes("show_status")) {
    db.exec("ALTER TABLE shows ADD COLUMN show_status TEXT NOT NULL DEFAULT 'confirmed'");
  }
  if (!showColumns.includes("google_event_id")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_event_id TEXT");
  }
  if (!showColumns.includes("google_sync_source")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_sync_source TEXT");
  }
  if (!showColumns.includes("google_sync_status")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_sync_status TEXT");
  }
  if (!showColumns.includes("google_notes")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_notes TEXT");
  }
  if (!showColumns.includes("google_last_synced_at")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_last_synced_at TEXT");
  }
  if (!showColumns.includes("needs_admin_completion")) {
    db.exec("ALTER TABLE shows ADD COLUMN needs_admin_completion INTEGER NOT NULL DEFAULT 0");
  }
  if (!showColumns.includes("google_archived")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_archived INTEGER NOT NULL DEFAULT 0");
  }
  if (!showColumns.includes("google_archived_at")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_archived_at TEXT");
  }
  if (!showColumns.includes("google_pinned")) {
    db.exec("ALTER TABLE shows ADD COLUMN google_pinned INTEGER NOT NULL DEFAULT 0");
  }
  if (!showColumns.includes("client_id")) {
    db.exec("ALTER TABLE shows ADD COLUMN client_id TEXT");
  }
  const clientColumns = db.prepare("PRAGMA table_info(clients)").all().map((column) => column.name);
  if (!clientColumns.includes("state")) {
    db.exec("ALTER TABLE clients ADD COLUMN state TEXT");
  }
  const invoiceColumns = db.prepare("PRAGMA table_info(invoices)").all().map((column) => column.name);
  if (!invoiceColumns.includes("client_id")) {
    db.exec("ALTER TABLE invoices ADD COLUMN client_id TEXT");
  }
  if (!invoiceColumns.includes("details_json")) {
    db.exec("ALTER TABLE invoices ADD COLUMN details_json TEXT NOT NULL DEFAULT '{}'");
  }
  const invoiceLineItemColumns = db.prepare("PRAGMA table_info(invoice_line_items)").all().map((column) => column.name);
  if (!invoiceLineItemColumns.includes("sac")) {
    db.exec("ALTER TABLE invoice_line_items ADD COLUMN sac TEXT");
  }
  if (!invoiceLineItemColumns.includes("custom_details")) {
    db.exec("ALTER TABLE invoice_line_items ADD COLUMN custom_details TEXT");
  }
  if (!invoiceLineItemColumns.includes("discount")) {
    db.exec("ALTER TABLE invoice_line_items ADD COLUMN discount TEXT");
  }
  if (!invoiceLineItemColumns.includes("discount_amount")) {
    db.exec("ALTER TABLE invoice_line_items ADD COLUMN discount_amount REAL NOT NULL DEFAULT 0");
  }
  db.exec(`
    UPDATE shows
    SET show_date_from = COALESCE(show_date_from, show_date),
        show_date_to = COALESCE(show_date_to, show_date),
        show_status = COALESCE(show_status, 'confirmed'),
        google_sync_status = COALESCE(google_sync_status, CASE WHEN google_event_id IS NOT NULL THEN 'synced' END),
        google_sync_source = COALESCE(google_sync_source, CASE WHEN google_event_id IS NOT NULL THEN 'pixelbug' END),
        needs_admin_completion = COALESCE(needs_admin_completion, 0),
        google_archived = COALESCE(google_archived, 0),
        google_pinned = COALESCE(google_pinned, 0)
    WHERE show_date_from IS NULL OR show_date_to IS NULL OR show_status IS NULL OR google_sync_status IS NULL OR google_sync_source IS NULL OR needs_admin_completion IS NULL OR google_archived IS NULL OR google_pinned IS NULL
  `);

  migrateLegacyJsonToDatabase();
  syncClientMasterFromExistingRecords(db);
  return db;
}

function roundMoney(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function getDiscountAmount(rawDiscount, baseAmount) {
  const normalized = String(rawDiscount || "").trim();
  if (!normalized) return 0;
  const numeric = Number(normalized.replace(/%$/, "").trim());
  if (!Number.isFinite(numeric) || numeric <= 0) return 0;
  const discountAmount = normalized.endsWith("%")
    ? Number(baseAmount || 0) * (numeric / 100)
    : numeric;
  return roundMoney(Math.min(Number(baseAmount || 0), Math.max(0, discountAmount)));
}

function isMaharashtraSupply(placeOfSupply) {
  return String(placeOfSupply || "").toLowerCase().includes("maharashtra");
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

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function getDueDateFromTerms(issueDate, paymentTerms) {
  const normalizedIssueDate = String(issueDate || "").trim();
  if (!normalizedIssueDate) return "";
  const normalizedTerms = String(paymentTerms || "Due on receipt").trim().toLowerCase();
  const offsets = {
    "due on receipt": 0,
    "net 10": 10,
    "net 15": 15,
    "net 30": 30
  };
  const offsetDays = offsets[normalizedTerms] ?? 0;
  return dateKey(addDays(parseDateKey(normalizedIssueDate), offsetDays));
}

function normalizeInvoiceLineItem(item, index = 0) {
  const quantity = Number(item.quantity || 0);
  const unitRate = Number(item.unitRate || 0);
  const grossAmount = roundMoney((quantity > 0 ? quantity : 1) * unitRate);
  const discount = String(item.discount || item.discountRaw || "").trim();
  const discountAmount = getDiscountAmount(discount, grossAmount);
  return {
    id: String(item.id || uid("line")).trim(),
    showId: String(item.showId || "").trim(),
    description: String(item.description || "").trim(),
    sac: String(item.sac || "").trim(),
    customDetails: String(item.customDetails || "").trim(),
    discount,
    discountAmount,
    quantity: quantity > 0 ? quantity : 1,
    unitRate: roundMoney(unitRate),
    amount: roundMoney(grossAmount - discountAmount),
    lineOrder: Number.isFinite(Number(item.lineOrder)) ? Number(item.lineOrder) : index
  };
}

function normalizeInvoice(invoice) {
  const lineItems = Array.isArray(invoice.lineItems)
    ? invoice.lineItems
        .map((item, index) => normalizeInvoiceLineItem(item, index))
        .filter((item) => item.description)
        .sort((a, b) => a.lineOrder - b.lineOrder)
    : [];
  const grossSubtotal = roundMoney(lineItems.reduce((sum, item) => sum + (item.quantity * item.unitRate), 0));
  const discountAmount = roundMoney(lineItems.reduce((sum, item) => sum + item.discountAmount, 0));
  const subtotal = roundMoney(grossSubtotal - discountAmount);
  const taxPercent = 18;
  const amountPaid = roundMoney(Math.max(0, Number(invoice.amountPaid || 0)));
  const rawDetails = invoice && typeof invoice.details === "object" && invoice.details !== null ? invoice.details : {};
  const details = {
    companyName: String(rawDetails.companyName || "PixelBug").trim(),
    companyAddress: String(rawDetails.companyAddress || "").trim(),
    companyEmail: String(rawDetails.companyEmail || "").trim(),
    companyPhone: String(rawDetails.companyPhone || "").trim(),
    companyGstin: String(rawDetails.companyGstin || "").trim(),
    clientBillingAddress: String(rawDetails.clientBillingAddress || "").trim(),
    clientGstin: String(rawDetails.clientGstin || "").trim(),
    placeOfSupply: String(rawDetails.placeOfSupply || "").trim(),
    paymentTerms: String(rawDetails.paymentTerms || "Due on receipt").trim(),
    bankAccountName: String(rawDetails.bankAccountName || "").trim(),
    bankName: String(rawDetails.bankName || "").trim(),
    bankAccountNumber: String(rawDetails.bankAccountNumber || "").trim(),
    bankIfsc: String(rawDetails.bankIfsc || "").trim(),
    footerNote: String(rawDetails.footerNote || "Please include the invoice number with your payment reference.").trim()
  };
  const intraState = isMaharashtraSupply(details.placeOfSupply);
  const sgstAmount = intraState ? roundMoney(subtotal * 0.09) : 0;
  const cgstAmount = intraState ? roundMoney(subtotal * 0.09) : 0;
  const igstAmount = intraState ? 0 : roundMoney(subtotal * 0.18);
  const taxAmount = roundMoney(sgstAmount + cgstAmount + igstAmount);
  const totalAmount = roundMoney(subtotal + taxAmount);
  const normalizedStatus = amountPaid >= totalAmount && totalAmount > 0
    ? "paid"
    : amountPaid > 0
      ? "partially_paid"
      : INVOICE_STATUSES.has(invoice.status)
        ? invoice.status
        : "draft";
  details.gstBreakup = {
    grossSubtotal,
    discountAmount,
    taxableAmount: subtotal,
    sgstAmount,
    cgstAmount,
    igstAmount,
    taxAmount
  };
  return {
    id: String(invoice.id || uid("invoice")).trim(),
    invoiceNumber: String(invoice.invoiceNumber || "").trim(),
    clientId: String(invoice.clientId || "").trim(),
    clientName: String(invoice.clientName || "").trim(),
    issueDate: String(invoice.issueDate || "").trim(),
    dueDate: String(invoice.dueDate || getDueDateFromTerms(invoice.issueDate, details.paymentTerms) || "").trim(),
    status: normalizedStatus,
    notes: String(invoice.notes || "").trim(),
    taxPercent: roundMoney(taxPercent),
    subtotal,
    taxAmount,
    totalAmount,
    amountPaid,
    balanceDue: roundMoney(Math.max(0, totalAmount - amountPaid)),
    currency: String(invoice.currency || "INR").trim() || "INR",
    details,
    createdAt: String(invoice.createdAt || new Date().toISOString()).trim(),
    updatedAt: new Date().toISOString(),
    lineItems,
    paymentEntries: Array.isArray(invoice.paymentEntries) ? invoice.paymentEntries : []
  };
}

function normalizeClient(client) {
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
    updatedAt: new Date().toISOString()
  };
}

function validateClients(clients) {
  const seenKeys = new Set();
  for (const rawClient of clients) {
    const client = normalizeClient(rawClient);
    if (!client.name) {
      return "Client name is required.";
    }
    const uniquenessKey = [
      client.name.toLowerCase(),
      client.state.toLowerCase(),
      client.gstin.toLowerCase()
    ].join("|");
    if (seenKeys.has(uniquenessKey)) {
      return "Client entries must remain unique by name, state, and GSTIN.";
    }
    seenKeys.add(uniquenessKey);
  }
  return "";
}

function validateInvoice(invoice, database = ensureDatabase()) {
  if (!invoice.invoiceNumber) {
    return "Invoice number is required.";
  }
  if (!invoice.clientId) {
    return "Client selection is required.";
  }
  const existingClient = database.prepare("SELECT id, name FROM clients WHERE id = ?").get(invoice.clientId);
  if (!existingClient) {
    return "Selected client was not found.";
  }
  if (!invoice.clientName || String(existingClient.name || "").trim() !== String(invoice.clientName || "").trim()) {
    return "Invoice client does not match the selected client.";
  }
  if (!invoice.issueDate) {
    return "Issue date is required.";
  }
  if (!invoice.lineItems.length) {
    return "At least one invoice line item is required.";
  }
  if (invoice.dueDate && invoice.issueDate && invoice.dueDate < invoice.issueDate) {
    return "Due date cannot be earlier than issue date.";
  }
  return "";
}

function readInvoices() {
  const database = ensureDatabase();
  const invoices = database.prepare(`
    SELECT id, invoice_number, client_name, client_id, issue_date, due_date, status, notes, details_json, tax_percent, subtotal, tax_amount, total_amount, amount_paid, currency, created_at, updated_at
    FROM invoices
    ORDER BY issue_date DESC, created_at DESC, invoice_number COLLATE NOCASE DESC
  `).all().map((invoice) => ({
    id: invoice.id,
    invoiceNumber: invoice.invoice_number,
    clientId: invoice.client_id || "",
    clientName: invoice.client_name,
    issueDate: invoice.issue_date,
    dueDate: invoice.due_date || "",
    status: invoice.status,
    notes: invoice.notes || "",
    details: (() => {
      try {
        const parsed = JSON.parse(invoice.details_json || "{}");
        return parsed && typeof parsed === "object" ? parsed : {};
      } catch (error) {
        return {};
      }
    })(),
    taxPercent: Number(invoice.tax_percent || 0),
    subtotal: Number(invoice.subtotal || 0),
    taxAmount: Number(invoice.tax_amount || 0),
    totalAmount: Number(invoice.total_amount || 0),
    amountPaid: Number(invoice.amount_paid || 0),
    balanceDue: roundMoney(Number(invoice.total_amount || 0) - Number(invoice.amount_paid || 0)),
    currency: invoice.currency || "INR",
    createdAt: invoice.created_at,
    updatedAt: invoice.updated_at,
    lineItems: [],
    paymentEntries: []
  }));

  const lineItemsByInvoiceId = new Map(invoices.map((invoice) => [invoice.id, invoice.lineItems]));
  const lineItems = database.prepare(`
    SELECT id, invoice_id, show_id, description, sac, custom_details, discount, discount_amount, quantity, unit_rate, amount, line_order
    FROM invoice_line_items
    ORDER BY invoice_id, line_order ASC, id ASC
  `).all();

  lineItems.forEach((item) => {
    const invoiceItems = lineItemsByInvoiceId.get(item.invoice_id);
    if (!invoiceItems) return;
    invoiceItems.push({
      id: item.id,
      showId: item.show_id || "",
      description: item.description,
      sac: item.sac || "",
      customDetails: item.custom_details || "",
      discount: item.discount || "",
      discountAmount: Number(item.discount_amount || 0),
      quantity: Number(item.quantity || 0),
      unitRate: Number(item.unit_rate || 0),
      amount: Number(item.amount || 0),
      lineOrder: Number(item.line_order || 0)
    });
  });

  const paymentsByInvoiceId = new Map(invoices.map((invoice) => [invoice.id, invoice.paymentEntries]));
  const payments = database.prepare(`
    SELECT id, invoice_id, payment_date, amount, note, created_at
    FROM invoice_payments
    ORDER BY payment_date ASC, created_at ASC, id ASC
  `).all();
  payments.forEach((payment) => {
    const invoicePayments = paymentsByInvoiceId.get(payment.invoice_id);
    if (!invoicePayments) return;
    invoicePayments.push({
      id: payment.id,
      paymentDate: payment.payment_date,
      amount: Number(payment.amount || 0),
      note: payment.note || "",
      createdAt: payment.created_at
    });
  });

  invoices.forEach((invoice) => {
    if (!invoice.paymentEntries.length) return;
    invoice.amountPaid = roundMoney(invoice.paymentEntries.reduce((sum, payment) => sum + Number(payment.amount || 0), 0));
    invoice.balanceDue = roundMoney(Math.max(0, Number(invoice.totalAmount || 0) - invoice.amountPaid));
    if (invoice.amountPaid >= Number(invoice.totalAmount || 0) && Number(invoice.totalAmount || 0) > 0) {
      invoice.status = "paid";
    } else if (invoice.amountPaid > 0) {
      invoice.status = "partially_paid";
    }
  });

  return invoices;
}

function saveInvoice(invoiceInput) {
  const database = ensureDatabase();
  const invoice = normalizeInvoice(invoiceInput);
  const validationError = validateInvoice(invoice, database);
  if (validationError) {
    throw new Error(validationError);
  }

  const existingInvoice = database.prepare("SELECT id, created_at FROM invoices WHERE id = ?").get(invoice.id);
  const duplicateInvoiceNumber = database.prepare("SELECT id FROM invoices WHERE invoice_number = ? AND id != ?").get(invoice.invoiceNumber, invoice.id);
  if (duplicateInvoiceNumber) {
    throw new Error("Invoice number already exists.");
  }

  database.exec("BEGIN");
  try {
    database.prepare(`
      INSERT OR REPLACE INTO invoices (
        id, invoice_number, client_name, client_id, issue_date, due_date, status, notes, details_json, tax_percent, subtotal, tax_amount, total_amount, amount_paid, currency, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      invoice.id,
      invoice.invoiceNumber,
      invoice.clientName,
      invoice.clientId || null,
      invoice.issueDate,
      invoice.dueDate || null,
      invoice.status,
      invoice.notes || null,
      JSON.stringify(invoice.details || {}),
      invoice.taxPercent,
      invoice.subtotal,
      invoice.taxAmount,
      invoice.totalAmount,
      invoice.amountPaid,
      invoice.currency,
      existingInvoice?.created_at || invoice.createdAt,
      invoice.updatedAt
    );

    database.prepare("DELETE FROM invoice_line_items WHERE invoice_id = ?").run(invoice.id);
    const insertLineItem = database.prepare(`
      INSERT INTO invoice_line_items (
        id, invoice_id, show_id, description, sac, custom_details, discount, discount_amount, quantity, unit_rate, amount, line_order
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    invoice.lineItems.forEach((item, index) => {
      insertLineItem.run(
        item.id,
        invoice.id,
        item.showId || null,
        item.description,
        item.sac || null,
        item.customDetails || null,
        item.discount || null,
        item.discountAmount || 0,
        item.quantity,
        item.unitRate,
        item.amount,
        index
      );
    });

    database.exec("COMMIT");
    return invoice.id;
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function addInvoicePayment(invoiceId, paymentInput = {}) {
  const database = ensureDatabase();
  const invoice = database.prepare("SELECT id, total_amount, amount_paid FROM invoices WHERE id = ?").get(invoiceId);
  if (!invoice) {
    throw new Error("Invoice not found.");
  }

  const paymentDate = String(paymentInput.paymentDate || "").trim();
  const amount = roundMoney(Number(paymentInput.amount || 0));
  const note = String(paymentInput.note || "").trim();
  if (!paymentDate) {
    throw new Error("Payment date is required.");
  }
  if (!Number.isFinite(amount) || amount <= 0) {
    throw new Error("Payment amount must be greater than zero.");
  }

  const existingPaymentCount = database.prepare("SELECT COUNT(*) AS count FROM invoice_payments WHERE invoice_id = ?").get(invoiceId)?.count || 0;
  const existingPaymentTotal = roundMoney((database.prepare("SELECT COALESCE(SUM(amount), 0) AS total FROM invoice_payments WHERE invoice_id = ?").get(invoiceId)?.total) || 0);
  const effectiveExistingPaid = existingPaymentCount ? existingPaymentTotal : roundMoney(Number(invoice.amount_paid || 0));
  const remainingBalance = roundMoney(Math.max(0, Number(invoice.total_amount || 0) - effectiveExistingPaid));
  if (amount > remainingBalance) {
    throw new Error(`Payment amount cannot exceed remaining balance of ${remainingBalance}.`);
  }
  const now = new Date().toISOString();
  database.exec("BEGIN");
  try {
    if (!existingPaymentCount && Number(invoice.amount_paid || 0) > 0) {
      database.prepare(`
        INSERT INTO invoice_payments (id, invoice_id, payment_date, amount, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
      `).run(uid("payment"), invoiceId, paymentDate, roundMoney(Number(invoice.amount_paid || 0)), "Legacy paid amount", now);
    }

    database.prepare(`
      INSERT INTO invoice_payments (id, invoice_id, payment_date, amount, note, created_at)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(uid("payment"), invoiceId, paymentDate, amount, note || null, now);

    const totalPaid = roundMoney((database.prepare("SELECT COALESCE(SUM(amount), 0) AS total FROM invoice_payments WHERE invoice_id = ?").get(invoiceId)?.total) || 0);
    const totalAmount = roundMoney(Number(invoice.total_amount || 0));
    const nextStatus = totalPaid >= totalAmount && totalAmount > 0 ? "paid" : totalPaid > 0 ? "partially_paid" : "sent";
    database.prepare("UPDATE invoices SET amount_paid = ?, status = ?, updated_at = ? WHERE id = ?").run(totalPaid, nextStatus, now, invoiceId);
    database.exec("COMMIT");
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function deleteInvoice(invoiceId) {
  const database = ensureDatabase();
  database.exec("BEGIN");
  try {
    database.prepare("DELETE FROM invoice_payments WHERE invoice_id = ?").run(invoiceId);
    database.prepare("DELETE FROM invoice_line_items WHERE invoice_id = ?").run(invoiceId);
    const result = database.prepare("DELETE FROM invoices WHERE id = ?").run(invoiceId);
    database.exec("COMMIT");
    return Number(result.changes || 0) > 0;
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function saveClient(clientInput) {
  const database = ensureDatabase();
  const client = normalizeClient(clientInput || {});
  const validationError = validateClients([client]);
  if (validationError) {
    throw new Error(validationError);
  }

  const duplicateClient = database.prepare(`
    SELECT id FROM clients
    WHERE lower(name) = lower(?)
      AND lower(COALESCE(state, '')) = lower(?)
      AND lower(COALESCE(gstin, '')) = lower(?)
      AND id != ?
  `).get(client.name, client.state || "", client.gstin || "", client.id);
  if (duplicateClient) {
    throw new Error("A client with the same name, state, and GSTIN already exists.");
  }

  const existingClient = database.prepare("SELECT id, created_at FROM clients WHERE id = ?").get(client.id);

  database.exec("BEGIN");
  try {
    database.prepare(`
      INSERT OR REPLACE INTO clients (
        id, name, state, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      client.id,
      client.name,
      client.state || null,
      client.billingAddress || null,
      client.gstin || null,
      client.contactName || null,
      client.contactEmail || null,
      client.contactPhone || null,
      client.notes || null,
      existingClient?.created_at || client.createdAt,
      client.updatedAt
    );

    database.prepare("UPDATE shows SET client = ? WHERE client_id = ?").run(client.name, client.id);
    database.prepare("UPDATE invoices SET client_name = ? WHERE client_id = ?").run(client.name, client.id);

    database.exec("COMMIT");
    return client.id;
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function deleteClient(clientId, options = {}) {
  const database = ensureDatabase();
  const linkedShow = database.prepare("SELECT id FROM shows WHERE client_id = ? LIMIT 1").get(clientId);
  const linkedInvoice = database.prepare("SELECT id FROM invoices WHERE client_id = ? LIMIT 1").get(clientId);
  if ((linkedShow || linkedInvoice) && !options.keepHistory) {
    throw new Error("This client is linked to shows or invoices and cannot be deleted yet.");
  }
  database.exec("BEGIN");
  try {
    if (options.keepHistory) {
      database.prepare("UPDATE shows SET client_id = NULL WHERE client_id = ?").run(clientId);
      database.prepare("UPDATE invoices SET client_id = NULL WHERE client_id = ?").run(clientId);
    }
    const result = database.prepare("DELETE FROM clients WHERE id = ?").run(clientId);
    database.exec("COMMIT");
    return Number(result.changes || 0) > 0;
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function getSetting(key) {
  const database = ensureDatabase();
  const row = database.prepare("SELECT value FROM app_settings WHERE key = ?").get(key);
  return row?.value ?? null;
}

function setSetting(key, value) {
  const database = ensureDatabase();
  if (value === null || value === undefined || value === "") {
    database.prepare("DELETE FROM app_settings WHERE key = ?").run(key);
    return;
  }
  database.prepare("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)").run(key, String(value));
}

function getGoogleTokens() {
  const raw = getSetting(GOOGLE_TOKEN_SETTING_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function setGoogleTokens(tokens) {
  setSetting(GOOGLE_TOKEN_SETTING_KEY, JSON.stringify(tokens || {}));
}

function readFirstConfiguredEnv(keys = []) {
  for (const key of keys) {
    const value = process.env[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return "";
}

function getGoogleConfig(req) {
  const clientId = readFirstConfiguredEnv(["GOOGLE_CLIENT_ID", "GOOGLE_OAUTH_CLIENT_ID", "GOOGLE_CLIENTID"]);
  const clientSecret = readFirstConfiguredEnv(["GOOGLE_CLIENT_SECRET", "GOOGLE_OAUTH_CLIENT_SECRET", "GOOGLE_SECRET"]);
  const calendarId = readFirstConfiguredEnv(["GOOGLE_CALENDAR_ID", "GOOGLE_SHARED_CALENDAR_ID", "CALENDAR_ID"]);
  const baseUrl = (process.env.PIXELBUG_BASE_URL || `http://${req?.headers?.host || `${HOST}:${PORT}`}`).trim().replace(/\/+$/, "");
  const missing = [];
  if (!clientId) missing.push("GOOGLE_CLIENT_ID");
  if (!clientSecret) missing.push("GOOGLE_CLIENT_SECRET");
  if (!calendarId) missing.push("GOOGLE_CALENDAR_ID");

  return {
    clientId,
    clientSecret,
    calendarId,
    baseUrl,
    redirectUri: `${baseUrl}/api/google/callback`,
    configured: missing.length === 0,
    missing
  };
}

function getGoogleStatus(req) {
  const config = getGoogleConfig(req);
  const tokens = getGoogleTokens();
  return {
    configured: config.configured,
    connected: Boolean(tokens?.refresh_token || tokens?.access_token),
    calendarId: config.calendarId || "",
    redirectUri: config.redirectUri,
    missingConfig: config.missing || [],
    lastSyncAt: getSetting(GOOGLE_LAST_SYNC_SETTING_KEY) || "",
    lastError: getSetting(GOOGLE_LAST_ERROR_SETTING_KEY) || ""
  };
}

function migrateLegacyJsonToDatabase() {
  const database = db;
  const counts = {
    users: database.prepare("SELECT COUNT(*) AS count FROM users").get().count,
    shows: database.prepare("SELECT COUNT(*) AS count FROM shows").get().count
  };
  if (counts.users || counts.shows || !fs.existsSync(DATA_FILE)) return;

  const parsed = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  const users = Array.isArray(parsed.users) ? parsed.users : [];
  const shows = Array.isArray(parsed.shows) ? parsed.shows : [];
  const clients = Array.isArray(parsed.clients) ? parsed.clients : [];
  const insertUser = database.prepare(`
    INSERT INTO users (
      id, name, email, phone, role, approved, color, email_verified, verification_token, reset_token, reset_token_expires_at, password_hash
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertShow = database.prepare(`
    INSERT INTO shows (
      id, show_date, show_date_from, show_date_to, show_status, google_event_id, google_sync_source, google_sync_status, google_notes, google_last_synced_at, needs_admin_completion, google_archived, google_archived_at, google_pinned, show_name, client_id, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertClient = database.prepare(`
    INSERT OR IGNORE INTO clients (
      id, name, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  database.exec("BEGIN");
  try {
    users.forEach((user) => {
      insertUser.run(
        user.id,
        user.name,
        String(user.email || "").toLowerCase(),
        user.phone || "",
        user.role || "crew",
        user.approved ? 1 : 0,
        user.color || null,
        user.emailVerified ?? 1,
        user.verificationToken ?? null,
        user.resetToken ?? null,
        user.resetTokenExpiresAt ?? null,
        user.passwordHash
      );
    });

    clients.forEach((client) => {
      const normalizedClient = normalizeClient(client);
      if (!normalizedClient.name) return;
      insertClient.run(
        normalizedClient.id,
        normalizedClient.name,
        normalizedClient.billingAddress || null,
        normalizedClient.gstin || null,
        normalizedClient.contactName || null,
        normalizedClient.contactEmail || null,
        normalizedClient.contactPhone || null,
        normalizedClient.notes || null,
        normalizedClient.createdAt,
        normalizedClient.updatedAt
      );
    });

    shows.forEach((show) => {
      const normalized = normalizeShow(show);
      insertShow.run(
        normalized.id,
        normalized.showDate,
        normalized.showDateFrom,
        normalized.showDateTo,
        normalized.showStatus,
        normalized.googleEventId,
        normalized.googleSyncSource,
        normalized.googleSyncStatus,
        normalized.googleNotes,
        normalized.googleLastSyncedAt,
        normalized.needsAdminCompletion ? 1 : 0,
        normalized.googleArchived ? 1 : 0,
        normalized.googleArchivedAt,
        normalized.googlePinned ? 1 : 0,
        normalized.showName,
        normalized.clientId || null,
        normalized.client,
        normalized.venue,
        normalized.location,
        normalized.showTime,
        normalized.amountShow,
        JSON.stringify(normalized.assignments)
      );
    });
    database.exec("COMMIT");
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function syncClientMasterFromExistingRecords(database) {
  const existingClients = database.prepare(`
    SELECT id, name, state, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
    FROM clients
    ORDER BY name COLLATE NOCASE
  `).all();
  const clientMap = new Map(existingClients.map((client) => [String(client.name || "").trim().toLowerCase(), client]));
  const insertClient = database.prepare(`
    INSERT INTO clients (
      id, name, state, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const linkShowClient = database.prepare("UPDATE shows SET client_id = ? WHERE id = ?");
  const linkInvoiceClient = database.prepare("UPDATE invoices SET client_id = ? WHERE id = ?");
  const now = new Date().toISOString();

  const ensureClient = (name) => {
    const normalizedName = String(name || "").trim();
    if (!normalizedName) return null;
    const key = normalizedName.toLowerCase();
    const existing = clientMap.get(key);
    if (existing) {
      return existing.id;
    }
    const client = {
      id: uid("client"),
      name: normalizedName,
      state: null,
      billing_address: null,
      gstin: null,
      contact_name: null,
      contact_email: null,
      contact_phone: null,
      notes: null,
      created_at: now,
      updated_at: now
    };
    insertClient.run(
      client.id,
      client.name,
      client.state,
      client.billing_address,
      client.gstin,
      client.contact_name,
      client.contact_email,
      client.contact_phone,
      client.notes,
      client.created_at,
      client.updated_at
    );
    clientMap.set(key, client);
    return client.id;
  };

  database.exec("BEGIN");
  try {
    database.prepare("SELECT id, client_id, client FROM shows").all().forEach((show) => {
      if (show.client_id) return;
      const clientId = ensureClient(show.client);
      if (clientId) {
        linkShowClient.run(clientId, show.id);
      }
    });
    database.prepare("SELECT id, client_id, client_name FROM invoices").all().forEach((invoice) => {
      if (invoice.client_id) return;
      const clientId = ensureClient(invoice.client_name);
      if (clientId) {
        linkInvoiceClient.run(clientId, invoice.id);
      }
    });
    database.exec("COMMIT");
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function readStore() {
  const database = ensureDatabase();
  const users = database.prepare(`
    SELECT id, name, email, phone, role, approved, color, email_verified, verification_token, reset_token, reset_token_expires_at, password_hash
    FROM users
    ORDER BY name COLLATE NOCASE
  `).all().map((user) => ({
    id: user.id,
    name: user.name,
    email: user.email,
    phone: user.phone,
    role: user.role,
    approved: Boolean(user.approved),
    color: user.color || null,
    emailVerified: Boolean(user.email_verified),
    verificationToken: user.verification_token || null,
    resetToken: user.reset_token || null,
    resetTokenExpiresAt: user.reset_token_expires_at || null,
    passwordHash: user.password_hash
  }));

  const shows = database.prepare(`
    SELECT id, show_date, show_date_from, show_date_to, show_status, show_name, client_id, client, venue, location, show_time, amount_show, assignments_json
    , google_event_id, google_sync_source, google_sync_status, google_notes, google_last_synced_at, needs_admin_completion, google_archived, google_archived_at, google_pinned
    FROM shows
    ORDER BY COALESCE(show_date_from, show_date) ASC, show_time ASC, show_name COLLATE NOCASE
  `).all().map((show) => ({
    id: show.id,
    showDate: show.show_date,
    showDateFrom: show.show_date_from || show.show_date,
    showDateTo: show.show_date_to || show.show_date_from || show.show_date,
    showStatus: show.show_status || "confirmed",
    googleEventId: show.google_event_id || "",
    googleSyncSource: show.google_sync_source || "",
    googleSyncStatus: show.google_sync_status || "",
    googleNotes: show.google_notes || "",
    googleLastSyncedAt: show.google_last_synced_at || "",
    needsAdminCompletion: Boolean(show.needs_admin_completion),
    googleArchived: Boolean(show.google_archived),
    googleArchivedAt: show.google_archived_at || "",
    googlePinned: Boolean(show.google_pinned),
    showName: show.show_name,
    clientId: show.client_id || "",
    client: show.client || "",
    venue: show.venue || "",
    location: show.location || "",
    showTime: show.show_time || "",
    amountShow: Number(show.amount_show || 0),
    assignments: JSON.parse(show.assignments_json || "[]")
  }));

  const clients = database.prepare(`
    SELECT id, name, state, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
    FROM clients
    ORDER BY name COLLATE NOCASE
  `).all().map((client) => ({
    id: client.id,
    name: client.name,
    state: client.state || "",
    billingAddress: client.billing_address || "",
    gstin: client.gstin || "",
    contactName: client.contact_name || "",
    contactEmail: client.contact_email || "",
    contactPhone: client.contact_phone || "",
    notes: client.notes || "",
    createdAt: client.created_at,
    updatedAt: client.updated_at
  }));

  return { users, shows, clients };
}

function writeStore(store) {
  const database = ensureDatabase();
  const replaceUser = database.prepare(`
    INSERT OR REPLACE INTO users (
      id, name, email, phone, role, approved, color, email_verified, verification_token, reset_token, reset_token_expires_at, password_hash
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const replaceShow = database.prepare(`
    INSERT OR REPLACE INTO shows (
      id, show_date, show_date_from, show_date_to, show_status, google_event_id, google_sync_source, google_sync_status, google_notes, google_last_synced_at, needs_admin_completion, google_archived, google_archived_at, google_pinned, show_name, client_id, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const replaceClient = database.prepare(`
    INSERT OR REPLACE INTO clients (
      id, name, state, billing_address, gstin, contact_name, contact_email, contact_phone, notes, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const clearUsers = database.prepare("DELETE FROM users");
  const clearShows = database.prepare("DELETE FROM shows");
  const clearClients = database.prepare("DELETE FROM clients");

  database.exec("BEGIN");
  try {
    clearUsers.run();
    clearShows.run();
    clearClients.run();

    store.users.forEach((user) => {
      replaceUser.run(
        user.id,
        user.name,
        String(user.email || "").toLowerCase(),
        user.phone || "",
        user.role,
        user.approved ? 1 : 0,
        user.color || null,
        user.emailVerified ? 1 : 0,
        user.verificationToken || null,
        user.resetToken || null,
        user.resetTokenExpiresAt || null,
        user.passwordHash
      );
    });

    (store.clients || []).forEach((client) => {
      const normalizedClient = normalizeClient(client);
      replaceClient.run(
        normalizedClient.id,
        normalizedClient.name,
        normalizedClient.state || null,
        normalizedClient.billingAddress || null,
        normalizedClient.gstin || null,
        normalizedClient.contactName || null,
        normalizedClient.contactEmail || null,
        normalizedClient.contactPhone || null,
        normalizedClient.notes || null,
        normalizedClient.createdAt,
        normalizedClient.updatedAt
      );
    });

    store.shows.forEach((show) => {
      const normalized = normalizeShow(show);
      replaceShow.run(
        normalized.id,
        normalized.showDate,
        normalized.showDateFrom,
        normalized.showDateTo,
        normalized.showStatus,
        normalized.googleEventId,
        normalized.googleSyncSource,
        normalized.googleSyncStatus,
        normalized.googleNotes,
        normalized.googleLastSyncedAt,
        normalized.needsAdminCompletion ? 1 : 0,
        normalized.googleArchived ? 1 : 0,
        normalized.googleArchivedAt,
        normalized.googlePinned ? 1 : 0,
        normalized.showName,
        normalized.clientId || null,
        normalized.client,
        normalized.venue,
        normalized.location,
        normalized.showTime,
        normalized.amountShow,
        JSON.stringify(normalized.assignments)
      );
    });
    database.exec("COMMIT");
  } catch (error) {
    database.exec("ROLLBACK");
    throw error;
  }
}

function parseCookies(req) {
  return (req.headers.cookie || "").split(";").reduce((cookies, entry) => {
    const [key, ...rest] = entry.trim().split("=");
    if (!key) return cookies;
    cookies[key] = decodeURIComponent(rest.join("="));
    return cookies;
  }, {});
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      if (!chunks.length) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch (error) {
        reject(new Error("Invalid JSON body."));
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, statusCode, payload, extraHeaders = {}) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    ...extraHeaders
  });
  res.end(JSON.stringify(payload));
}

function sanitizeUser(user) {
  return {
    id: user.id,
    name: user.name,
    email: user.email,
    phone: user.phone,
    role: user.role,
    approved: Boolean(user.approved),
    color: user.color || null,
    emailVerified: Boolean(user.emailVerified)
  };
}

function getSessionUser(req, store) {
  const cookies = parseCookies(req);
  const sessionId = cookies[SESSION_COOKIE];
  const userId = sessionId ? sessions.get(sessionId) : null;
  if (!userId) return null;
  return store.users.find((user) => user.id === userId) || null;
}

function createSession(userId) {
  const sessionId = crypto.randomBytes(24).toString("hex");
  sessions.set(sessionId, userId);
  return sessionId;
}

function buildSessionCookie(value, maxAge = null) {
  const parts = [
    `${SESSION_COOKIE}=${value}`,
    "HttpOnly",
    "Path=/",
    "SameSite=Strict"
  ];

  if (COOKIE_SECURE) {
    parts.push("Secure");
  }

  if (maxAge !== null) {
    parts.push(`Max-Age=${maxAge}`);
  }

  return parts.join("; ");
}

function clearSession(req, res) {
  const cookies = parseCookies(req);
  const sessionId = cookies[SESSION_COOKIE];
  if (sessionId) {
    sessions.delete(sessionId);
  }
}

function hashPassword(password, salt = crypto.randomBytes(16).toString("hex")) {
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

function isStrongPassword(password) {
  return password.length >= 8
    && /[a-z]/.test(password)
    && /[A-Z]/.test(password)
    && /\d/.test(password);
}

function verifyPassword(password, storedValue) {
  if (!storedValue || !storedValue.includes(":")) return false;
  const [salt, originalHash] = storedValue.split(":");
  const candidateHash = crypto.scryptSync(password, salt, 64);
  const originalBuffer = Buffer.from(originalHash, "hex");
  if (candidateHash.length !== originalBuffer.length) return false;
  return crypto.timingSafeEqual(candidateHash, originalBuffer);
}

function normalizeShow(show) {
  const showDateFrom = String(show.showDateFrom || show.showDate || "");
  const rawShowDateTo = String(show.showDateTo || show.showDateFrom || show.showDate || "");
  const showDateTo = rawShowDateTo && showDateFrom && rawShowDateTo < showDateFrom
    ? showDateFrom
    : rawShowDateTo;
  return {
    id: show.id || uid("show"),
    showDate: showDateFrom,
    showDateFrom,
    showDateTo,
    showStatus: show.showStatus === "tentative" ? "tentative" : "confirmed",
    googleEventId: String(show.googleEventId || "").trim(),
    googleSyncSource: String(show.googleSyncSource || "").trim(),
    googleSyncStatus: String(show.googleSyncStatus || "").trim(),
    googleNotes: String(show.googleNotes || "").trim(),
    googleLastSyncedAt: String(show.googleLastSyncedAt || "").trim(),
    needsAdminCompletion: Boolean(show.needsAdminCompletion),
    googleArchived: Boolean(show.googleArchived),
    googleArchivedAt: String(show.googleArchivedAt || "").trim(),
    googlePinned: Boolean(show.googlePinned),
    showName: String(show.showName || "").trim(),
    clientId: String(show.clientId || "").trim(),
    client: String(show.client || "").trim(),
    venue: String(show.venue || "").trim(),
    location: String(show.location || "").trim(),
    showTime: String(show.showTime || ""),
    amountShow: Number(show.amountShow || 0),
    assignments: Array.isArray(show.assignments)
      ? show.assignments
          .filter((assignment) => assignment && (assignment.crewId || assignment.manualCrewName))
          .map((assignment) => ({
            crewId: String(assignment.crewId || ""),
            manualCrewName: String(assignment.manualCrewName || "").trim(),
            lightDesignerId: String(assignment.lightDesignerId || ""),
            operatorAmount: Number(assignment.operatorAmount || 0),
            onwardTravelDate: String(assignment.onwardTravelDate || assignment.travelDate || ""),
            returnTravelDate: String(assignment.returnTravelDate || ""),
            onwardTravelSector: String(assignment.onwardTravelSector || assignment.travelSector || "").trim(),
            returnTravelSector: String(assignment.returnTravelSector || "").trim(),
            notes: String(assignment.notes || assignment.travelNotes || "").trim()
          }))
      : []
  };
}

function getCrewSummaryFromShow(show, store) {
  return (show.assignments || [])
    .map((assignment) => store.users.find((user) => user.id === assignment.crewId)?.name || assignment.manualCrewName)
    .filter(Boolean)
    .join(", ");
}

function hasGoogleSyncedFieldChanges(previousShow, nextShow, store) {
  if (!previousShow) {
    return shouldSyncShowWithGoogle(nextShow);
  }
  return (
    String(previousShow.showName || "") !== String(nextShow.showName || "") ||
    String(getShowStartDate(previousShow) || "") !== String(getShowStartDate(nextShow) || "") ||
    String(getShowEndDate(previousShow) || "") !== String(getShowEndDate(nextShow) || "") ||
    String(previousShow.location || "") !== String(nextShow.location || "") ||
    String(previousShow.googleNotes || "") !== String(nextShow.googleNotes || "") ||
    getCrewSummaryFromShow(previousShow, store) !== getCrewSummaryFromShow(nextShow, store)
  );
}

function hasApprovedAdmin(store) {
  return store.users.some((user) => user.role === "admin" && user.approved);
}

function canAccessInvoices(user) {
  return user?.role === "admin" || user?.role === "accounts";
}

function isAdminRouteAllowedForUser(pathname, user) {
  if (user?.role === "admin") return true;
  if (canAccessInvoices(user) && (pathname === "/api/admin/invoices" || pathname.startsWith("/api/admin/invoices/"))) {
    return true;
  }
  return false;
}

function usedApprovedColors(users, excludeUserId = null) {
  return users
    .filter((user) => (user.role === "crew" || user.role === "admin") && user.approved && user.id !== excludeUserId)
    .map((user) => user.color)
    .filter(Boolean);
}

function validateUniqueEmails(users) {
  const seen = new Set();
  for (const user of users) {
    const email = String(user.email || "").toLowerCase();
    if (!email || seen.has(email)) {
      return false;
    }
    seen.add(email);
  }
  return true;
}

function validateCrewColors(users) {
  const seen = new Set();
  for (const user of users) {
    if (!(user.role === "crew" || user.role === "admin") || !user.approved || !user.color) continue;
    if (seen.has(user.color)) {
      return false;
    }
    seen.add(user.color);
  }
  return true;
}

function buildVerificationPath(token) {
  return `/index.html?verify=${encodeURIComponent(token)}`;
}

function buildResetPath(token) {
  return `/index.html?reset=${encodeURIComponent(token)}`;
}

async function sendEmail({ to, subject, html, text }) {
  ensureDataStore();
  const resendKey = process.env.RESEND_API_KEY;
  const fromEmail = process.env.PIXELBUG_EMAIL_FROM;

  if (resendKey && fromEmail) {
    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${resendKey}`
      },
      body: JSON.stringify({
        from: fromEmail,
        to: [to],
        subject,
        html,
        text
      })
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Email delivery failed: ${body || response.statusText}`);
    }

    return { delivered: true, provider: "resend" };
  }

  const entry = {
    createdAt: new Date().toISOString(),
    to,
    subject,
    html,
    text
  };
  fs.appendFileSync(OUTBOX_FILE, `${JSON.stringify(entry)}\n`);
  return { delivered: false, provider: "outbox", outboxFile: OUTBOX_FILE };
}

async function deliverResetEmail(user) {
  const resetPath = buildResetPath(user.resetToken);
  const absoluteUrl = `${process.env.PIXELBUG_BASE_URL || `http://${HOST}:${PORT}`}${resetPath}`;
  const delivery = await sendEmail({
    to: user.email,
    subject: "Reset your PixelBug password",
    html: `<p>Reset your PixelBug password by opening this link:</p><p><a href="${absoluteUrl}">${absoluteUrl}</a></p>`,
    text: `Reset your PixelBug password: ${absoluteUrl}`
  });
  return { resetPath, absoluteUrl, delivery };
}

function formatDateKeyFromDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10);
}

function getShowStartDate(show) {
  return String(show.showDateFrom || show.showDate || "");
}

function getShowEndDate(show) {
  return String(show.showDateTo || show.showDateFrom || show.showDate || "");
}

function parseDateKey(dateStr) {
  const match = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function toDateKey(date) {
  return new Date(date).toISOString().slice(0, 10);
}

function shiftDateKey(dateStr, days) {
  const date = parseDateKey(dateStr);
  if (!date) return "";
  date.setUTCDate(date.getUTCDate() + days);
  return toDateKey(date);
}

function getGoogleSyncStartDateKey() {
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1))
    .toISOString()
    .slice(0, 10);
}

function shouldSyncShowWithGoogle(show) {
  const showEnd = getShowEndDate(show);
  if (!showEnd) return false;
  return showEnd >= getGoogleSyncStartDateKey();
}

function shouldPushShowToGoogle(show) {
  if (!shouldSyncShowWithGoogle(show)) return false;
  if (!getShowStartDate(show) || !getShowEndDate(show)) return false;
  return !show.googleEventId || show.googleSyncStatus === "pending_push" || show.googleSyncStatus === "sync_error";
}

function parseGoogleDescription(description = "") {
  const text = String(description || "").trim();
  const crewMatch = text.match(/(?:^|\n)Crew:\s*(.+?)(?:\n|$)/i);
  const notesMatch = text.match(/(?:^|\n)Notes:\s*([\s\S]*)$/i);
  return {
    crewNames: crewMatch
      ? crewMatch[1].split(",").map((item) => item.trim()).filter(Boolean)
      : [],
    notes: notesMatch ? notesMatch[1].trim() : ""
  };
}

function buildGoogleDescription(show, store) {
  const crewNames = (show.assignments || [])
    .map((assignment) => store.users.find((user) => user.id === assignment.crewId)?.name || assignment.manualCrewName)
    .filter(Boolean)
    .join(", ");
  const notes = String(show.googleNotes || "").trim();
  const lines = [];
  if (crewNames) lines.push(`Crew: ${crewNames}`);
  if (notes) lines.push(`Notes: ${notes}`);
  return lines.join("\n\n");
}

function mapGoogleEventToShow(event, store, existingShow = null) {
  const isAllDay = Boolean(event.start?.date);
  const rawShowDateFrom = isAllDay
    ? String(event.start.date || "")
    : formatDateKeyFromDateTime(event.start?.dateTime);
  const rawShowDateTo = isAllDay
    ? shiftDateKey(String(event.end?.date || event.start?.date || ""), -1)
    : formatDateKeyFromDateTime(event.end?.dateTime || event.start?.dateTime);
  const showDateFrom = rawShowDateFrom;
  const showDateTo = rawShowDateTo && showDateFrom && rawShowDateTo < showDateFrom
    ? showDateFrom
    : rawShowDateTo;
  const parsedDescription = parseGoogleDescription(event.description || "");
  const crewUsers = store.users.filter((user) => (user.role === "crew" || user.role === "admin") && user.approved);
  const matchedCrewIds = parsedDescription.crewNames
    .map((name) => crewUsers.find((user) => user.name.toLowerCase() === name.toLowerCase())?.id)
    .filter(Boolean);
  const previousAssignments = new Map((existingShow?.assignments || []).map((assignment) => [assignment.crewId, assignment]));
  const assignments = matchedCrewIds.map((crewId) => {
    const existingAssignment = previousAssignments.get(crewId);
    return existingAssignment || {
      crewId,
      lightDesignerId: "",
      operatorAmount: 0,
      onwardTravelDate: "",
      returnTravelDate: "",
      onwardTravelSector: "",
      returnTravelSector: "",
      notes: ""
    };
  });

  if (existingShow?.googleSyncStatus === "pending_push") {
    return normalizeShow({
      ...existingShow,
      googleSyncSource: "pixelbug",
      googleNotes: parsedDescription.notes || existingShow.googleNotes || "",
      googleLastSyncedAt: existingShow.googleLastSyncedAt || ""
    });
  }

  if (existingShow?.googleSyncStatus === "sync_error") {
    return normalizeShow({
      ...existingShow,
      googleSyncSource: "pixelbug",
      googleNotes: parsedDescription.notes || existingShow.googleNotes || "",
      googleLastSyncedAt: existingShow.googleLastSyncedAt || ""
    });
  }

  return normalizeShow({
    ...(existingShow || {}),
    id: existingShow?.id || uid("show"),
    showDateFrom,
    showDateTo,
    showDate: showDateFrom,
    showName: event.summary || existingShow?.showName || "Untitled Google Event",
    location: event.location || existingShow?.location || "",
    googleEventId: event.id || existingShow?.googleEventId || "",
    googleSyncSource: "google",
    googleSyncStatus: existingShow?.googleEventId ? "updated_from_google" : "synced",
    googleNotes: parsedDescription.notes,
    googleLastSyncedAt: new Date().toISOString(),
    needsAdminCompletion: existingShow?.needsAdminCompletion || !existingShow?.client || Number(existingShow?.amountShow || 0) === 0,
    googleArchived: existingShow?.googleArchived || false,
    googleArchivedAt: existingShow?.googleArchivedAt || "",
    googlePinned: existingShow?.googlePinned || false,
    assignments
  });
}

async function exchangeGoogleCode(req, code) {
  const config = getGoogleConfig(req);
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: config.clientId,
      client_secret: config.clientSecret,
      redirect_uri: config.redirectUri,
      grant_type: "authorization_code"
    }).toString()
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || "Google token exchange failed.");
  }
  return payload;
}

async function refreshGoogleAccessToken(req, tokens) {
  const config = getGoogleConfig(req);
  if (!tokens?.refresh_token) {
    throw new Error("Google refresh token is missing.");
  }

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      refresh_token: tokens.refresh_token,
      grant_type: "refresh_token"
    }).toString()
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || "Google token refresh failed.");
  }
  return {
    ...tokens,
    access_token: payload.access_token,
    expires_at: Date.now() + (Number(payload.expires_in || 3600) * 1000),
    token_type: payload.token_type || tokens.token_type || "Bearer"
  };
}

async function getGoogleAccessToken(req) {
  const config = getGoogleConfig(req);
  if (!config.configured) {
    throw new Error("Google Calendar integration is not configured.");
  }

  let tokens = getGoogleTokens();
  if (!tokens) {
    throw new Error("Google Calendar is not connected.");
  }

  if (!tokens.access_token || !tokens.expires_at || Date.now() > (tokens.expires_at - 60_000)) {
    tokens = await refreshGoogleAccessToken(req, tokens);
    setGoogleTokens(tokens);
  }

  return tokens.access_token;
}

async function googleApiRequest(req, method, apiPath, body = null, query = null) {
  const accessToken = await getGoogleAccessToken(req);
  const url = new URL(`https://www.googleapis.com/calendar/v3${apiPath}`);
  if (query) {
    Object.entries(query).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== "") {
        url.searchParams.set(key, String(value));
      }
    });
  }

  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error?.message || payload.error_description || "Google Calendar request failed.");
  }
  return payload;
}

async function pullGoogleCalendarIntoStore(req, currentStore) {
  const config = getGoogleConfig(req);
  const timeMin = new Date(`${getGoogleSyncStartDateKey()}T00:00:00.000Z`);
  const timeMax = new Date();
  timeMax.setMonth(timeMax.getMonth() + 18);

  const payload = await googleApiRequest(
    req,
    "GET",
    `/calendars/${encodeURIComponent(config.calendarId)}/events`,
    null,
    {
      singleEvents: false,
      showDeleted: true,
      maxResults: 1000,
      timeMin: timeMin.toISOString(),
      timeMax: timeMax.toISOString()
    }
  );

  const nextStore = {
    users: [...currentStore.users],
    shows: [...currentStore.shows],
    clients: [...(currentStore.clients || [])]
  };
  const showsByGoogleId = new Map(nextStore.shows.filter((show) => show.googleEventId).map((show) => [show.googleEventId, show]));
  const cancelledGoogleIds = new Set();

  (payload.items || []).forEach((event) => {
    if (event.status === "cancelled") {
      if (event.id) cancelledGoogleIds.add(event.id);
      return;
    }
    const existingShow = showsByGoogleId.get(event.id) || null;
    const mappedShow = mapGoogleEventToShow(event, nextStore, existingShow);
    if (existingShow) {
      const index = nextStore.shows.findIndex((show) => show.id === existingShow.id);
      nextStore.shows[index] = mappedShow;
    } else {
      nextStore.shows.push(mappedShow);
    }
  });

  if (cancelledGoogleIds.size) {
    nextStore.shows = nextStore.shows.filter((show) => {
      if (!show.googleEventId || !cancelledGoogleIds.has(show.googleEventId)) return true;
      return show.googleSyncStatus === "pending_push";
    });
  }

  setSetting(GOOGLE_LAST_SYNC_SETTING_KEY, new Date().toISOString());
  return nextStore;
}

async function pushPixelbugShowsToGoogle(req, currentStore) {
  const config = getGoogleConfig(req);
  const nextStore = {
    users: [...currentStore.users],
    shows: [...currentStore.shows],
    clients: [...(currentStore.clients || [])]
  };
  const syncErrors = [];

  for (const show of nextStore.shows) {
    if (!shouldPushShowToGoogle(show)) continue;
    const startDate = getShowStartDate(show);
    const endDateExclusive = shiftDateKey(getShowEndDate(show), 1);
    if (!startDate || !endDateExclusive || endDateExclusive <= startDate) {
      show.googleSyncSource = "pixelbug";
      show.googleSyncStatus = "sync_error";
      syncErrors.push(`${show.showName}: Invalid show date range for Google sync.`);
      continue;
    }
    try {
      const body = {
        summary: show.showName,
        location: show.location || undefined,
        description: buildGoogleDescription(show, nextStore) || undefined,
        start: { date: startDate },
        end: { date: endDateExclusive }
      };

      const syncedEvent = show.googleEventId
        ? await googleApiRequest(req, "PATCH", `/calendars/${encodeURIComponent(config.calendarId)}/events/${encodeURIComponent(show.googleEventId)}`, body)
        : await googleApiRequest(req, "POST", `/calendars/${encodeURIComponent(config.calendarId)}/events`, body);

      show.googleEventId = syncedEvent.id;
      show.googleSyncSource = "pixelbug";
      show.googleSyncStatus = "synced";
      show.googleNotes = parseGoogleDescription(syncedEvent.description || "").notes || show.googleNotes || "";
      show.googleLastSyncedAt = new Date().toISOString();
    } catch (error) {
      show.googleSyncSource = "pixelbug";
      show.googleSyncStatus = "sync_error";
      syncErrors.push(`${show.showName}: ${error.message || "Google push failed."}`);
    }
  }

  setSetting(GOOGLE_LAST_SYNC_SETTING_KEY, new Date().toISOString());
  setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, syncErrors[0] || "");
  return nextStore;
}

async function maybeAutoPullGoogleCalendar(req, currentUser, store) {
  if (!currentUser || currentUser.role !== "admin") return store;
  const status = getGoogleStatus(req);
  if (!status.configured || !status.connected) return store;
  const lastSync = status.lastSyncAt ? Date.parse(status.lastSyncAt) : 0;
  if (lastSync && (Date.now() - lastSync) < GOOGLE_SYNC_DEBOUNCE_MS) return store;

  try {
    let workingStore = store;
    if (workingStore.shows.some((show) => shouldPushShowToGoogle(show))) {
      workingStore = await pushPixelbugShowsToGoogle(req, workingStore);
    }
    const syncedStore = await pullGoogleCalendarIntoStore(req, workingStore);
    writeStore(syncedStore);
    return syncedStore;
  } catch (error) {
    setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, error.message || "Google sync failed.");
    return store;
  }
}

function sendBootstrap(res, store, currentUser, extraHeaders = {}, req = null) {
  sendJson(res, 200, {
    users: store.users.map(sanitizeUser),
    shows: store.shows,
    clients: store.clients || [],
    invoices: canAccessInvoices(currentUser) ? readInvoices() : [],
    currentUserId: currentUser?.id || null,
    hasAdmin: hasApprovedAdmin(store),
    google: getGoogleStatus(req)
  }, extraHeaders);
}

async function handleApi(req, res) {
  let store = readStore();
  let currentUser = getSessionUser(req, store);
  const requestUrl = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`);
  const pathname = requestUrl.pathname;

  if (req.method === "GET" && pathname === "/api/bootstrap") {
    store = await maybeAutoPullGoogleCalendar(req, currentUser, store);
    currentUser = getSessionUser(req, store);
    sendBootstrap(res, store, currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/setup-admin") {
    if (hasApprovedAdmin(store)) {
      sendJson(res, 409, { error: "Admin already exists." });
      return true;
    }

    const body = await readJson(req);
    const name = String(body.name || "").trim();
    const email = String(body.email || "").trim().toLowerCase();
    const phone = String(body.phone || "").trim();
    const password = String(body.password || "");
    const color = String(body.color || "").trim();

    if (!name || !email || !phone || !color) {
      sendJson(res, 400, { error: "All admin fields are required." });
      return true;
    }

    if (!isStrongPassword(password)) {
      sendJson(res, 400, { error: "Password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    const admin = {
      id: uid("admin"),
      name,
      email,
      phone,
      role: "admin",
      approved: true,
      color,
      emailVerified: true,
      passwordHash: hashPassword(password)
    };

    store.users.push(admin);
    writeStore(store);
    const sessionId = createSession(admin.id);
    sendBootstrap(res, store, admin, {
      "Set-Cookie": buildSessionCookie(sessionId)
    }, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/login") {
    const body = await readJson(req);
    const email = String(body.email || "").trim().toLowerCase();
    const password = String(body.password || "");
    const user = store.users.find((item) => item.email.toLowerCase() === email);

    if (!user || !verifyPassword(password, user.passwordHash)) {
      sendJson(res, 401, { error: "Invalid email or password." });
      return true;
    }

    if (!user.approved) {
      sendJson(res, 403, { error: "Account request is pending admin approval." });
      return true;
    }

    const sessionId = createSession(user.id);
    sendBootstrap(res, store, user, {
      "Set-Cookie": buildSessionCookie(sessionId)
    }, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/logout") {
    clearSession(req, res);
    sendJson(res, 200, { ok: true }, {
      "Set-Cookie": buildSessionCookie("", 0)
    });
    return true;
  }

  if (req.method === "POST" && pathname === "/api/register") {
    const body = await readJson(req);
    const role = String(body.role || "");
    const email = String(body.email || "").trim().toLowerCase();
    const color = body.color ? String(body.color).trim() : null;

    if (!SELF_SERVICE_ROLES.has(role)) {
      sendJson(res, 400, { error: "Invalid account type." });
      return true;
    }

    if (store.users.some((user) => user.email.toLowerCase() === email)) {
      sendJson(res, 409, { error: "That email already exists." });
      return true;
    }

    if (!isStrongPassword(String(body.password || ""))) {
      sendJson(res, 400, { error: "Password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    if ((role === "crew" || role === "admin") && (!color || usedApprovedColors(store.users).includes(color))) {
      sendJson(res, 409, { error: "That crew color is not available." });
      return true;
    }

    const newUser = {
      id: uid(role),
      name: String(body.name || "").trim(),
      email,
      phone: String(body.phone || "").trim(),
      role,
      approved: false,
      color: role === "crew" || role === "admin" ? color : null,
      emailVerified: true,
      verificationToken: null,
      passwordHash: hashPassword(String(body.password || ""))
    };

    store.users.push(newUser);

    writeStore(store);
    sendJson(res, 200, {
      ok: true
    });
    return true;
  }

  if (req.method === "POST" && pathname === "/api/request-password-reset") {
    const body = await readJson(req);
    const email = String(body.email || "").trim().toLowerCase();
    const user = store.users.find((item) => item.email.toLowerCase() === email);

    if (!user) {
      sendJson(res, 200, { ok: true });
      return true;
    }

    user.resetToken = crypto.randomBytes(24).toString("hex");
    user.resetTokenExpiresAt = Date.now() + (1000 * 60 * 30);
    writeStore(store);
    const delivery = await deliverResetEmail(user);
    sendJson(res, 200, {
      ok: true,
      resetPath: delivery.resetPath,
      resetUrl: delivery.absoluteUrl,
      emailDelivery: delivery.delivery
    });
    return true;
  }

  if (req.method === "POST" && pathname === "/api/reset-password") {
    const body = await readJson(req);
    const token = String(body.token || "");
    const newPassword = String(body.newPassword || "");
    const user = store.users.find((item) => item.resetToken === token);

    if (!user || !user.resetTokenExpiresAt || user.resetTokenExpiresAt < Date.now()) {
      sendJson(res, 400, { error: "Reset link is invalid or expired." });
      return true;
    }

    if (!isStrongPassword(newPassword)) {
      sendJson(res, 400, { error: "New password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    user.passwordHash = hashPassword(newPassword);
    user.resetToken = null;
    user.resetTokenExpiresAt = null;
    writeStore(store);
    sendJson(res, 200, { ok: true });
    return true;
  }

  if (pathname.startsWith("/api/admin/") && !isAdminRouteAllowedForUser(pathname, currentUser)) {
    if (!currentUser) {
      sendJson(res, 403, { error: "Login required." });
      return true;
    }
    if (currentUser.role === "accounts") {
      sendJson(res, 403, { error: "Accounts access is limited to invoicing." });
      return true;
    }
    sendJson(res, 403, { error: "Admin access required." });
    return true;
  }

  if (req.method === "GET" && pathname === "/api/admin/google/status") {
    sendJson(res, 200, getGoogleStatus(req));
    return true;
  }

  if (req.method === "GET" && pathname === "/api/admin/google/auth") {
    const config = getGoogleConfig(req);
    if (!config.configured) {
      sendJson(res, 400, { error: "Google Calendar env vars are not configured." });
      return true;
    }

    const stateToken = crypto.randomBytes(20).toString("hex");
    googleOauthStates.set(stateToken, {
      userId: currentUser.id,
      createdAt: Date.now()
    });
    const authUrl = new URL("https://accounts.google.com/o/oauth2/v2/auth");
    authUrl.searchParams.set("client_id", config.clientId);
    authUrl.searchParams.set("redirect_uri", config.redirectUri);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("scope", "https://www.googleapis.com/auth/calendar");
    authUrl.searchParams.set("access_type", "offline");
    authUrl.searchParams.set("prompt", "consent");
    authUrl.searchParams.set("state", stateToken);
    sendJson(res, 200, { authUrl: authUrl.toString() });
    return true;
  }

  if (req.method === "GET" && pathname === "/api/google/callback") {
    const url = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`);
    const code = url.searchParams.get("code");
    const stateToken = url.searchParams.get("state");
    const stateEntry = stateToken ? googleOauthStates.get(stateToken) : null;

    if (!code || !stateEntry || !store.users.find((user) => user.id === stateEntry.userId && user.role === "admin")) {
      res.writeHead(302, { Location: "/index.html?google=error" });
      res.end();
      return true;
    }

    try {
      const tokens = await exchangeGoogleCode(req, code);
      setGoogleTokens({
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token,
        token_type: tokens.token_type || "Bearer",
        expires_at: Date.now() + (Number(tokens.expires_in || 3600) * 1000)
      });
      setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, "");
      googleOauthStates.delete(stateToken);
      res.writeHead(302, { Location: "/index.html?google=connected" });
      res.end();
    } catch (error) {
      setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, error.message || "Google auth failed.");
      res.writeHead(302, { Location: "/index.html?google=error" });
      res.end();
    }
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/google/sync") {
    try {
      const pushedStore = await pushPixelbugShowsToGoogle(req, store);
      const syncedStore = await pullGoogleCalendarIntoStore(req, pushedStore);
      writeStore(syncedStore);
      sendBootstrap(res, syncedStore, currentUser, {}, req);
    } catch (error) {
      setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, error.message || "Google sync failed.");
      sendJson(res, 400, { error: error.message || "Google sync failed." });
    }
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/add-crew") {
    const body = await readJson(req);
    const name = String(body.name || "").trim();
    const email = String(body.email || "").trim().toLowerCase();
    const phone = String(body.phone || "").trim();
    const password = String(body.password || "");
    const color = String(body.color || "").trim();

    if (!name || !email || !phone || !color) {
      sendJson(res, 400, { error: "All crew fields are required." });
      return true;
    }

    if (!isStrongPassword(password)) {
      sendJson(res, 400, { error: "Password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    if (store.users.some((user) => user.email.toLowerCase() === email)) {
      sendJson(res, 409, { error: "That email already exists." });
      return true;
    }

    if (usedApprovedColors(store.users).includes(color)) {
      sendJson(res, 409, { error: "That crew color is not available." });
      return true;
    }

    const newUser = {
      id: uid("crew"),
      name,
      email,
      phone,
      role: "crew",
      approved: true,
      color,
      emailVerified: true,
      verificationToken: null,
      passwordHash: hashPassword(password)
    };
    store.users.push(newUser);

    writeStore(store);
    sendBootstrap(res, store, currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/add-viewer") {
    const body = await readJson(req);
    const name = String(body.name || "").trim();
    const email = String(body.email || "").trim().toLowerCase();
    const phone = String(body.phone || "").trim();
    const password = String(body.password || "");

    if (!name || !email || !phone) {
      sendJson(res, 400, { error: "All view-only fields are required." });
      return true;
    }

    if (!isStrongPassword(password)) {
      sendJson(res, 400, { error: "Password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    if (store.users.some((user) => user.email.toLowerCase() === email)) {
      sendJson(res, 409, { error: "That email already exists." });
      return true;
    }

    const newUser = {
      id: uid("viewer"),
      name,
      email,
      phone,
      role: "viewer",
      approved: true,
      color: null,
      emailVerified: true,
      verificationToken: null,
      passwordHash: hashPassword(password)
    };
    store.users.push(newUser);

    writeStore(store);
    sendBootstrap(res, store, currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/add-accounts") {
    const body = await readJson(req);
    const name = String(body.name || "").trim();
    const email = String(body.email || "").trim().toLowerCase();
    const phone = String(body.phone || "").trim();
    const password = String(body.password || "");

    if (!name || !email || !phone) {
      sendJson(res, 400, { error: "All accounts fields are required." });
      return true;
    }

    if (!isStrongPassword(password)) {
      sendJson(res, 400, { error: "Password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    if (store.users.some((user) => user.email.toLowerCase() === email)) {
      sendJson(res, 409, { error: "That email already exists." });
      return true;
    }

    const newUser = {
      id: uid("accounts"),
      name,
      email,
      phone,
      role: "accounts",
      approved: true,
      color: null,
      emailVerified: true,
      verificationToken: null,
      passwordHash: hashPassword(password)
    };
    store.users.push(newUser);

    writeStore(store);
    sendBootstrap(res, store, currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/change-password") {
    if (!currentUser) {
      sendJson(res, 403, { error: "Login required." });
      return true;
    }

    const body = await readJson(req);
    const currentPassword = String(body.currentPassword || "");
    const nextPassword = String(body.newPassword || "");

    const storedUser = store.users.find((user) => user.id === currentUser.id);
    if (!storedUser || !verifyPassword(currentPassword, storedUser.passwordHash)) {
      sendJson(res, 401, { error: "Current password is incorrect." });
      return true;
    }

    if (!isStrongPassword(nextPassword)) {
      sendJson(res, 400, { error: "New password must be 8+ characters and include upper, lower, and number." });
      return true;
    }

    storedUser.passwordHash = hashPassword(nextPassword);
    writeStore(store);
    sendJson(res, 200, { ok: true });
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/invoices") {
    const body = await readJson(req);
    saveInvoice(body || {});
    sendBootstrap(res, readStore(), currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname.startsWith("/api/admin/invoices/") && pathname.endsWith("/payments")) {
    const invoiceId = decodeURIComponent(pathname.slice("/api/admin/invoices/".length, -"/payments".length));
    if (!invoiceId) {
      sendJson(res, 400, { error: "Invoice id is required." });
      return true;
    }
    const body = await readJson(req);
    addInvoicePayment(invoiceId, body || {});
    sendBootstrap(res, readStore(), currentUser, {}, req);
    return true;
  }

  if (req.method === "DELETE" && pathname.startsWith("/api/admin/invoices/")) {
    const invoiceId = decodeURIComponent(pathname.slice("/api/admin/invoices/".length));
    if (!invoiceId) {
      sendJson(res, 400, { error: "Invoice id is required." });
      return true;
    }
    const deleted = deleteInvoice(invoiceId);
    if (!deleted) {
      sendJson(res, 404, { error: "Invoice not found." });
      return true;
    }
    sendBootstrap(res, readStore(), currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/gstin-lookup") {
    const body = await readJson(req);
    const result = await lookupGstinDetails(body?.gstin || "");
    sendJson(res, 200, result);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/clients") {
    const body = await readJson(req);
    saveClient(body || {});
    sendBootstrap(res, readStore(), currentUser, {}, req);
    return true;
  }

  if (req.method === "DELETE" && pathname.startsWith("/api/admin/clients/")) {
    const clientId = decodeURIComponent(pathname.slice("/api/admin/clients/".length));
    if (!clientId) {
      sendJson(res, 400, { error: "Client id is required." });
      return true;
    }
    const deleted = deleteClient(clientId, { keepHistory: requestUrl.searchParams.get("keepHistory") === "true" });
    if (!deleted) {
      sendJson(res, 404, { error: "Client not found." });
      return true;
    }
    sendBootstrap(res, readStore(), currentUser, {}, req);
    return true;
  }

  if (req.method === "POST" && pathname === "/api/admin/state") {
    const body = await readJson(req);
    const incomingUsers = Array.isArray(body.users) ? body.users : [];
    const incomingShows = Array.isArray(body.shows) ? body.shows : [];
    const incomingClients = Array.isArray(body.clients) ? body.clients : [];

    const usersById = new Map(store.users.map((user) => [user.id, user]));
    const nextUsers = [];

    for (const item of incomingUsers) {
      const existing = usersById.get(item.id);
      if (!existing) {
        sendJson(res, 400, { error: "Unknown user in update payload." });
        return true;
      }

      nextUsers.push({
        ...existing,
        name: String(item.name || "").trim(),
        email: String(item.email || "").trim().toLowerCase(),
        phone: String(item.phone || "").trim(),
        role: item.role === "admin" ? "admin" : item.role === "viewer" ? "viewer" : item.role === "accounts" ? "accounts" : "crew",
        approved: Boolean(item.approved),
        color: item.color || null
      });
    }

    if (!validateUniqueEmails(nextUsers)) {
      sendJson(res, 400, { error: "Duplicate email detected." });
      return true;
    }

    if (!validateCrewColors(nextUsers)) {
      sendJson(res, 400, { error: "Crew colors must remain unique." });
      return true;
    }

    if (!nextUsers.some((user) => user.role === "admin" && user.approved)) {
      sendJson(res, 400, { error: "At least one approved admin is required." });
      return true;
    }

    const nextClients = incomingClients.map(normalizeClient);
    const clientValidationError = validateClients(nextClients);
    if (clientValidationError) {
      sendJson(res, 400, { error: clientValidationError });
      return true;
    }

    const allowedUserIds = new Set(nextUsers.filter((user) => user.role === "crew" || user.role === "admin").map((user) => user.id));
    const clientNamesById = new Map(nextClients.map((client) => [client.id, client.name]));
    const clientIdsByName = new Map(nextClients.map((client) => [String(client.name || "").trim().toLowerCase(), client.id]));
    const existingShowsById = new Map(store.shows.map((show) => [show.id, show]));
    const clientIds = new Set(nextClients.map((client) => client.id));
    const nextShows = incomingShows.map(normalizeShow).map((show) => {
      const previousShow = existingShowsById.get(show.id);
      const resolvedClientId = show.clientId
        || clientIdsByName.get(String(show.client || "").trim().toLowerCase())
        || (previousShow?.clientId && clientIds.has(previousShow.clientId) ? previousShow.clientId : "");
      const filteredShow = {
        ...show,
        clientId: resolvedClientId,
        client: resolvedClientId ? (clientNamesById.get(resolvedClientId) || show.client) : show.client,
        assignments: show.assignments.filter((assignment) => allowedUserIds.has(assignment.crewId) || assignment.manualCrewName)
      };
      const isNewGoogleEligibleShow = !previousShow && shouldSyncShowWithGoogle(filteredShow);
      const hasPendingLocalSyncState = filteredShow.googleSyncStatus === "pending_push" || filteredShow.googleSyncStatus === "sync_error";
      const syncedFieldsChanged = filteredShow.googleEventId && hasGoogleSyncedFieldChanges(previousShow, filteredShow, { users: nextUsers });
      if (
        isNewGoogleEligibleShow ||
        hasPendingLocalSyncState ||
        syncedFieldsChanged
      ) {
        filteredShow.googleSyncSource = "pixelbug";
        filteredShow.googleSyncStatus = "pending_push";
      }
      return filteredShow;
    });

    const invalidShow = nextShows.find((show) => {
      if (show.clientId && clientIds.has(show.clientId)) return false;
      const previousShow = existingShowsById.get(show.id);
      const isLegacyBlankClientShow = previousShow && !show.clientId && !String(show.client || "").trim();
      return !isLegacyBlankClientShow;
    });
    if (invalidShow) {
      sendJson(res, 400, { error: `Show "${invalidShow.showName || "Untitled Show"}" must use a client from the client master.` });
      return true;
    }

    const stillExists = nextUsers.find((user) => user.id === currentUser.id);
    if (!stillExists || stillExists.role !== "admin" || !stillExists.approved) {
      sendJson(res, 400, { error: "Current admin account must remain approved." });
      return true;
    }

    let nextStore = { users: nextUsers, shows: nextShows, clients: nextClients };
    writeStore(nextStore);

    const googleStatus = getGoogleStatus(req);
    if (googleStatus.configured && googleStatus.connected) {
      try {
        nextStore = await pushPixelbugShowsToGoogle(req, nextStore);
        nextStore = await pullGoogleCalendarIntoStore(req, nextStore);
        writeStore(nextStore);
      } catch (error) {
        setSetting(GOOGLE_LAST_ERROR_SETTING_KEY, error.message || "Google push failed.");
      }
    }

    sendBootstrap(res, nextStore, stillExists, {}, req);
    return true;
  }

  return false;
}

function sendFile(filePath, res) {
  fs.readFile(filePath, (error, data) => {
    if (error) {
      res.writeHead(error.code === "ENOENT" ? 404 : 500, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(error.code === "ENOENT" ? "Not found" : "Server error");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
      "Cache-Control": "no-store"
    });
    res.end(data);
  });
}

const server = http.createServer(async (req, res) => {
  try {
    const handled = await handleApi(req, res);
    if (handled) return;

    const requestPath = req.url === "/" ? "/index.html" : req.url.split("?")[0];
    const safePath = path.normalize(requestPath).replace(/^(\.\.[/\\])+/, "");
    const filePath = path.join(ROOT, safePath);

    if (!filePath.startsWith(ROOT) || filePath.startsWith(DATA_DIR) || filePath === __filename) {
      res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Forbidden");
      return;
    }

    sendFile(filePath, res);
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Server error." });
  }
});

server.listen(PORT, HOST, () => {
  ensureDataStore();
  console.log(`PixelBug preview server running at http://${HOST}:${PORT}`);
});
