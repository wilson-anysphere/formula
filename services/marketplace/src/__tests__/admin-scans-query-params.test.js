const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const { createRequire } = require("node:module");
const os = require("node:os");
const path = require("node:path");

const { createMarketplaceServer } = require("../server");

const requireFromHere = createRequire(__filename);

function listen(server) {
  return new Promise((resolve, reject) => {
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      if (!addr || typeof addr === "string") {
        reject(new Error("Unexpected server.address() value"));
        return;
      }
      resolve(addr.port);
    });
    server.on("error", reject);
  });
}

function closeServer(server) {
  return new Promise((resolve) => {
    if (!server.listening) {
      resolve();
      return;
    }
    server.close(resolve);
  });
}

test("admin scans endpoint trims query params before filtering scans", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-admin-scans-"));
  const dataDir = path.join(tmpRoot, "data");
  const adminToken = "admin-secret-token";

  const { server, store } = await createMarketplaceServer({ dataDir, adminToken });
  let called = null;

  try {
    store.listPackageScans = async (opts) => {
      called = opts;
      return [];
    };

    const port = await listen(server);
    const baseUrl = `http://127.0.0.1:${port}`;

    const res = await fetch(
      `${baseUrl}/api/admin/scans?status=${encodeURIComponent("  passed  ")}&publisher=${encodeURIComponent(
        "  acme  ",
      )}&extensionId=${encodeURIComponent("  acme.test  ")}&limit=10&offset=5`,
      {
        headers: { Authorization: `Bearer ${adminToken}` },
      },
    );
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.deepEqual(body, { scans: [] });
    assert.deepEqual(called, {
      status: "passed",
      publisher: "acme",
      extensionId: "acme.test",
      limit: 10,
      offset: 5,
    });
  } finally {
    await closeServer(server);
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

