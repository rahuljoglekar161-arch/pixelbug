const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { DatabaseSync } = require("node:sqlite");

const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "store.json");
const DB_FILE = path.join(DATA_DIR, "pixelbug.db");
const OUTBOX_FILE = path.join(DATA_DIR, "email-outbox.log");
const SESSION_COOKIE = "pixelbug_session";
const sessions = new Map();
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
      show_name TEXT NOT NULL,
      client TEXT,
      venue TEXT,
      location TEXT,
      show_time TEXT,
      amount_show REAL NOT NULL DEFAULT 0,
      assignments_json TEXT NOT NULL DEFAULT '[]'
    );
  `);

  migrateLegacyJsonToDatabase();
  return db;
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
      id, show_date, show_name, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    SELECT id, show_date, show_name, client, venue, location, show_time, amount_show, assignments_json
    FROM shows
    ORDER BY show_date ASC, show_time ASC, show_name COLLATE NOCASE
  `).all().map((show) => ({
    id: show.id,
    showDate: show.show_date,
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
      id, show_date, show_name, client, venue, location, show_time, amount_show, assignments_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
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
  return {
    id: show.id || uid("show"),
    showDate: String(show.showDate || ""),
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
            operatorAmount: Number(assignment.operatorAmount || 0),
            travelDate: String(assignment.travelDate || ""),
            travelSector: String(assignment.travelSector || "").trim(),
            travelNotes: String(assignment.travelNotes || "").trim()
          }))
      : []
  };
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

async function deliverVerificationEmail(user) {
  const verificationPath = buildVerificationPath(user.verificationToken);
  const absoluteUrl = `${process.env.PIXELBUG_BASE_URL || `http://${HOST}:${PORT}`}${verificationPath}`;
  const delivery = await sendEmail({
    to: user.email,
    subject: "Verify your PixelBug account",
    html: `<p>Verify your PixelBug account by opening this link:</p><p><a href="${absoluteUrl}">${absoluteUrl}</a></p>`,
    text: `Verify your PixelBug account: ${absoluteUrl}`
  });
  return { verificationPath, absoluteUrl, delivery };
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

function sendBootstrap(res, store, currentUser, extraHeaders = {}) {
  sendJson(res, 200, {
    users: store.users.map(sanitizeUser),
    shows: store.shows,
    currentUserId: currentUser?.id || null,
    hasAdmin: hasApprovedAdmin(store)
  }, extraHeaders);
}

async function handleApi(req, res) {
  const store = readStore();
  const currentUser = getSessionUser(req, store);
  const pathname = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`).pathname;

  if (req.method === "GET" && pathname === "/api/bootstrap") {
    sendBootstrap(res, store, currentUser);
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
    });
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

    if (!user.emailVerified) {
      sendJson(res, 403, { error: "Verify your email before login." });
      return true;
    }

    if (!user.approved) {
      sendJson(res, 403, { error: "Account request is pending admin approval." });
      return true;
    }

    const sessionId = createSession(user.id);
    sendBootstrap(res, store, user, {
      "Set-Cookie": buildSessionCookie(sessionId)
    });
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

    const verificationToken = crypto.randomBytes(24).toString("hex");

    const newUser = {
      id: uid(role),
      name: String(body.name || "").trim(),
      email,
      phone: String(body.phone || "").trim(),
      role,
      approved: false,
      color: role === "crew" || role === "admin" ? color : null,
      emailVerified: false,
      verificationToken,
      passwordHash: hashPassword(String(body.password || ""))
    };

    store.users.push(newUser);

    writeStore(store);
    const delivery = await deliverVerificationEmail(newUser);
    sendJson(res, 200, {
      ok: true,
      verificationPath: delivery.verificationPath,
      verificationUrl: delivery.absoluteUrl,
      emailDelivery: delivery.delivery
    });
    return true;
  }

  if (req.method === "POST" && pathname === "/api/verify-email") {
    const body = await readJson(req);
    const token = String(body.token || "");
    const user = store.users.find((item) => item.verificationToken === token);
    if (!user) {
      sendJson(res, 404, { error: "Verification link is invalid or expired." });
      return true;
    }

    user.emailVerified = true;
    user.verificationToken = null;
    writeStore(store);
    sendJson(res, 200, { ok: true });
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

    const verificationToken = crypto.randomBytes(24).toString("hex");
    const newUser = {
      id: uid("crew"),
      name,
      email,
      phone,
      role: "crew",
      approved: true,
      color,
      emailVerified: false,
      verificationToken,
      passwordHash: hashPassword(password)
    };
    store.users.push(newUser);

    writeStore(store);
    const delivery = await deliverVerificationEmail(newUser);
    sendJson(res, 200, {
      ...{
        users: store.users.map(sanitizeUser),
        shows: store.shows,
        currentUserId: currentUser?.id || null,
        hasAdmin: hasApprovedAdmin(store)
      },
      verificationPath: delivery.verificationPath,
      verificationUrl: delivery.absoluteUrl,
      emailDelivery: delivery.delivery
    });
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
    const nextShows = incomingShows.map(normalizeShow).map((show) => ({
      ...show,
      assignments: show.assignments.filter((assignment) => allowedUserIds.has(assignment.crewId))
    }));

    const stillExists = nextUsers.find((user) => user.id === currentUser.id);
    if (!stillExists || stillExists.role !== "admin" || !stillExists.approved) {
      sendJson(res, 400, { error: "Current admin account must remain approved." });
      return true;
    }

    writeStore({ users: nextUsers, shows: nextShows });
    sendBootstrap(res, { users: nextUsers, shows: nextShows }, stillExists);
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
