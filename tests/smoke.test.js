const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const net = require("node:net");
const { spawn } = require("node:child_process");

function getFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      server.close((error) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(address.port);
      });
    });
    server.on("error", reject);
  });
}

async function waitForServer(baseUrl, child) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15000) {
    if (child.exitCode !== null) {
      throw new Error(`Server exited early with code ${child.exitCode}`);
    }
    try {
      const response = await fetch(`${baseUrl}/api/bootstrap`);
      if (response.ok) return;
    } catch (_) {
      // retry
    }
    await new Promise((resolve) => setTimeout(resolve, 150));
  }
  throw new Error("Server did not become ready in time.");
}

function createCookieJar() {
  let cookie = "";
  return {
    header() {
      return cookie;
    },
    absorb(response) {
      const setCookie = response.headers.get("set-cookie");
      if (!setCookie) return;
      cookie = setCookie.split(";")[0];
    }
  };
}

async function apiRequest(baseUrl, jar, pathname, options = {}) {
  const response = await fetch(`${baseUrl}${pathname}`, {
    method: options.method || "GET",
    headers: {
      "Content-Type": "application/json",
      ...(jar.header() ? { cookie: jar.header() } : {}),
      ...(options.headers || {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
    redirect: "manual"
  });
  jar.absorb(response);
  const payload = await response.json().catch(() => ({}));
  return { response, payload };
}

test("accounts users can manage invoices but not admin-only routes", async (t) => {
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), "pixelbug-smoke-"));
  const port = await getFreePort();
  const baseUrl = `http://127.0.0.1:${port}`;
  const child = spawn(process.execPath, ["server.js"], {
    cwd: path.resolve(__dirname, ".."),
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      PIXELBUG_DATA_DIR: tmpRoot,
      PIXELBUG_STORE_PATH: path.join(tmpRoot, "store.json"),
      PIXELBUG_DB_PATH: path.join(tmpRoot, "pixelbug.db"),
      PIXELBUG_OUTBOX_PATH: path.join(tmpRoot, "email-outbox.log")
    },
    stdio: "pipe"
  });

  let stderr = "";
  child.stderr.on("data", (chunk) => {
    stderr += chunk.toString("utf8");
  });

  await waitForServer(baseUrl, child);

  t.after(async () => {
    if (child.exitCode === null) {
      child.kill("SIGTERM");
      await new Promise((resolve) => child.once("exit", resolve));
    }
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  });

  const adminJar = createCookieJar();
  const accountsJar = createCookieJar();
  const clientId = "client_test_acme";

  {
    const { response } = await apiRequest(baseUrl, adminJar, "/api/setup-admin", {
      method: "POST",
      body: {
        name: "Admin User",
        email: "admin@example.com",
        phone: "9999999999",
        password: "AdminPass1",
        color: "#4285f4"
      }
    });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response } = await apiRequest(baseUrl, adminJar, "/api/admin/add-accounts", {
      method: "POST",
      body: {
        name: "Accounts User",
        email: "accounts@example.com",
        phone: "8888888888",
        password: "AccountsPass1"
      }
    });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response: bootstrapResponse, payload: bootstrapPayload } = await apiRequest(baseUrl, adminJar, "/api/bootstrap");
    assert.equal(bootstrapResponse.status, 200, stderr);
    const { response } = await apiRequest(baseUrl, adminJar, "/api/admin/state", {
      method: "POST",
      body: {
        users: bootstrapPayload.users,
        shows: bootstrapPayload.shows,
        clients: [
          {
            id: clientId,
            name: "Acme Corp",
            billingAddress: "Acme Corp\nMumbai",
            gstin: "27ABCDE1234F1Z5",
            contactName: "Acme Accounts",
            contactEmail: "accounts@acme.example",
            contactPhone: "9999999999",
            notes: ""
          }
        ]
      }
    });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response } = await apiRequest(baseUrl, adminJar, "/api/admin/invoices", {
      method: "POST",
      body: {
        invoiceNumber: "INV-TEST-001",
        clientId,
        clientName: "Acme Corp",
        issueDate: "2026-04-05",
        dueDate: "2026-04-15",
        status: "draft",
        taxPercent: 18,
        amountPaid: 0,
        notes: "Smoke test invoice",
        details: {
          companyName: "PixelBug LLP",
          clientBillingAddress: "Acme Corp\nMumbai",
          bankName: "Axis Bank",
          paymentTerms: "50% advance, balance within 7 days"
        },
        lineItems: [
          {
            description: "Console programming",
            quantity: 1,
            unitRate: 25000,
            lineOrder: 0
          }
        ]
      }
    });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response } = await apiRequest(baseUrl, adminJar, "/api/logout", { method: "POST" });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response } = await apiRequest(baseUrl, accountsJar, "/api/login", {
      method: "POST",
      body: {
        email: "accounts@example.com",
        password: "AccountsPass1"
      }
    });
    assert.equal(response.status, 200, stderr);
  }

  {
    const { response, payload } = await apiRequest(baseUrl, accountsJar, "/api/bootstrap");
    assert.equal(response.status, 200, stderr);
    assert.ok(Array.isArray(payload.invoices));
    assert.equal(payload.invoices.length, 1);
    assert.equal(payload.invoices[0].invoiceNumber, "INV-TEST-001");
    assert.equal(payload.invoices[0].clientId, clientId);
    assert.equal(payload.invoices[0].details.companyName, "PixelBug LLP");
    assert.equal(payload.invoices[0].details.bankName, "Axis Bank");
  }

  {
    const { response, payload } = await apiRequest(baseUrl, accountsJar, "/api/admin/invoices", {
      method: "POST",
      body: {
        invoiceNumber: "INV-TEST-002",
        clientId,
        clientName: "Acme Corp",
        issueDate: "2026-04-06",
        dueDate: "2026-04-16",
        status: "sent",
        taxPercent: 0,
        amountPaid: 0,
        notes: "",
        lineItems: [
          {
            description: "Follow-up billing",
            quantity: 1,
            unitRate: 5000,
            lineOrder: 0
          }
        ]
      }
    });
    assert.equal(response.status, 200, stderr);
    assert.ok(payload.invoices.some((invoice) => invoice.invoiceNumber === "INV-TEST-002"));
  }

  {
    const { payload: bootstrapPayload } = await apiRequest(baseUrl, accountsJar, "/api/bootstrap");
    const invoice = bootstrapPayload.invoices.find((item) => item.invoiceNumber === "INV-TEST-002");
    assert.ok(invoice);
    const { response, payload } = await apiRequest(baseUrl, accountsJar, `/api/admin/invoices/${encodeURIComponent(invoice.id)}/payments`, {
      method: "POST",
      body: {
        paymentDate: "2026-04-08",
        amount: 1000,
        note: "Advance"
      }
    });
    assert.equal(response.status, 200, stderr);
    const updatedInvoice = payload.invoices.find((item) => item.id === invoice.id);
    assert.equal(updatedInvoice.amountPaid, 1000);
    assert.equal(updatedInvoice.status, "partially_paid");
    assert.equal(updatedInvoice.paymentEntries.length, 1);
    assert.equal(updatedInvoice.paymentEntries[0].paymentDate, "2026-04-08");
  }

  {
    const { response, payload } = await apiRequest(baseUrl, accountsJar, "/api/admin/google/status");
    assert.equal(response.status, 403, stderr);
    assert.equal(payload.error, "Accounts access is limited to calendar view, shows view, clients, and invoicing.");
  }
});
