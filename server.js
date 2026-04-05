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
const DATA_DIR = path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "store.json");
const DB_FILE = path.join(DATA_DIR, "pixelbug.db");
const OUTBOX_FILE = path.join(DATA_DIR, "email-outbox.log");
const SESSION_COOKIE = "pixelbug_session";
const GOOGLE_TOKEN_SETTING_KEY = "google_oauth_tokens";
const GOOGLE_LAST_SYNC_SETTING_KEY = "google_last_sync_at";
const GOOGLE_LAST_ERROR_SETTING_KEY = "google_last_sync_error";
const GOOGLE_SYNC_DEBOUNCE_MS = 1000 * 30;
const sessions = new Map();
const googleOauthStates = new Map();
let db;
const IS_PRODUCTION = process.env.NODE_ENV === "production";
const COOKIE_SECURE = process.env.PIXELBUG_COOKIE_SECURE === "true" || IS_PRODUCTION;

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

function uid(prefix) {
  return `${prefix}_${crypto.randomBytes(6).toString("hex")}`;
}

function ensureDataStore() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify({ users: [], shows: [] }, null, 2));
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
      client TEXT,
      venue TEXT,
      location TEXT,
      show_time TEXT,
      amount_show REAL NOT NULL DEFAULT 0,
      assignments_json TEXT NOT NULL DEFAULT '[]'
    );
    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `);

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
  return db;
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

function getGoogleConfig(req) {
  const clientId = process.env.GOOGLE_CLIENT_ID || "";
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET || "";
  const calendarId = process.env.GOOGLE_CALENDAR_ID || "";
  const baseUrl = process.env.PIXELBUG_BASE_URL || `http://${req?.headers?.host || `${HOST}:${PORT}`}`;
  return {
    clientId,
    clientSecret,
    calendarId,
    baseUrl,
    redirectUri: `${baseUrl}/api/google/callback`,
    configured: Boolean(clientId && clientSecret && calendarId)
  };
}

function getGoogleStatus(req) {
  const config = getGoogleConfig(req);
  const tokens = getGoogleTokens();
  return {
    configured: config.configured,
    connected: Boolean(tokens?.refresh_token || tokens?.access_token),
    calendarId: config.calendarId || "",
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
  const insertUser = database.prepare(`
    INSERT INTO users (
      id, name, email, phone, role, approved, color, email_verified, verification_token, reset_token, reset_token_expires_at, password_hash
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertShow = database.prepare(`
    INSERT INTO shows (
      id, show_date, show_date_from, show_date_to, show_status, google_event_id, google_sync_source, google_sync_status, google_notes, google_last_synced_at, needs_admin_completion, google_archived, google_archived_at, google_pinned, show_name, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    SELECT id, show_date, show_date_from, show_date_to, show_status, show_name, client, venue, location, show_time, amount_show, assignments_json
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
    client: show.client || "",
    venue: show.venue || "",
    location: show.location || "",
    showTime: show.show_time || "",
    amountShow: Number(show.amount_show || 0),
    assignments: JSON.parse(show.assignments_json || "[]")
  }));

  return { users, shows };
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
      id, show_date, show_date_from, show_date_to, show_status, google_event_id, google_sync_source, google_sync_status, google_notes, google_last_synced_at, needs_admin_completion, google_archived, google_archived_at, google_pinned, show_name, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const clearUsers = database.prepare("DELETE FROM users");
  const clearShows = database.prepare("DELETE FROM shows");

  database.exec("BEGIN");
  try {
    clearUsers.run();
    clearShows.run();

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
    client: String(show.client || "").trim(),
    venue: String(show.venue || "").trim(),
    location: String(show.location || "").trim(),
    showTime: String(show.showTime || ""),
    amountShow: Number(show.amountShow || 0),
    assignments: Array.isArray(show.assignments)
      ? show.assignments
          .filter((assignment) => assignment && assignment.crewId)
          .map((assignment) => ({
            crewId: String(assignment.crewId),
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
    .map((assignment) => store.users.find((user) => user.id === assignment.crewId)?.name)
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
    .map((assignment) => store.users.find((user) => user.id === assignment.crewId)?.name)
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
    shows: [...currentStore.shows]
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
    shows: [...currentStore.shows]
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
    currentUserId: currentUser?.id || null,
    hasAdmin: hasApprovedAdmin(store),
    google: getGoogleStatus(req)
  }, extraHeaders);
}

async function handleApi(req, res) {
  let store = readStore();
  let currentUser = getSessionUser(req, store);
  const pathname = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`).pathname;

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

    if (!["admin", "crew", "viewer"].includes(role)) {
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

  if (!currentUser || currentUser.role !== "admin") {
    if (pathname.startsWith("/api/admin/")) {
      sendJson(res, 403, { error: "Admin access required." });
      return true;
    }
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

  if (req.method === "POST" && pathname === "/api/admin/state") {
    const body = await readJson(req);
    const incomingUsers = Array.isArray(body.users) ? body.users : [];
    const incomingShows = Array.isArray(body.shows) ? body.shows : [];

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
        role: item.role === "admin" ? "admin" : item.role === "viewer" ? "viewer" : "crew",
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

    const allowedUserIds = new Set(nextUsers.filter((user) => user.role === "crew" || user.role === "admin").map((user) => user.id));
    const existingShowsById = new Map(store.shows.map((show) => [show.id, show]));
    const nextShows = incomingShows.map(normalizeShow).map((show) => {
      const filteredShow = {
        ...show,
        assignments: show.assignments.filter((assignment) => allowedUserIds.has(assignment.crewId))
      };
      const previousShow = existingShowsById.get(filteredShow.id);
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

    const stillExists = nextUsers.find((user) => user.id === currentUser.id);
    if (!stillExists || stillExists.role !== "admin" || !stillExists.approved) {
      sendJson(res, 400, { error: "Current admin account must remain approved." });
      return true;
    }

    let nextStore = { users: nextUsers, shows: nextShows };
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
