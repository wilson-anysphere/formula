const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const { createRequire } = require("node:module");
const os = require("node:os");
const path = require("node:path");

const { createExtensionPackageV2 } = require("../../../../shared/extension-package");
const { generateEd25519KeyPair } = require("../../../../shared/crypto/signing");
const { createMarketplaceServer } = require("../server");

const requireFromHere = createRequire(__filename);

async function createTempExtensionDir({ publisher, name, version, jsSource }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-cors-"));
  await fs.mkdir(path.join(dir, "dist"), { recursive: true });

  const manifest = {
    name,
    publisher,
    version,
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(dir, "dist", "extension.js"), jsSource, "utf8");
  await fs.writeFile(path.join(dir, "README.md"), "hello\n");

  return { dir, manifest };
}

function parseHeaderList(value) {
  return new Set(
    String(value || "")
      .split(",")
      .map((v) => v.trim().toLowerCase())
      .filter(Boolean),
  );
}

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

test("public GET endpoints send CORS headers and expose signature/sha headers", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-cors-data-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "cors-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  const keys = generateEd25519KeyPair();
  const { server, store } = await createMarketplaceServer({ dataDir });

  try {
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    const port = await listen(server);
    const baseUrl = `http://127.0.0.1:${port}`;
    const origin = "tauri://localhost";

    // Search
    {
      const res = await fetch(`${baseUrl}/api/search?q=${encodeURIComponent(published.id)}`, {
        headers: { Origin: origin },
      });
      assert.equal(res.status, 200);
      assert.equal(res.headers.get("access-control-allow-origin"), "*");
      const exposed = parseHeaderList(res.headers.get("access-control-expose-headers"));
      assert.ok(exposed.has("etag"));
      assert.ok(exposed.has("retry-after"));
      assert.ok(exposed.has("x-package-sha256"));
      assert.ok(exposed.has("x-package-signature"));
      assert.ok(exposed.has("x-package-format-version"));
      assert.ok(exposed.has("x-publisher"));
      assert.ok(exposed.has("x-publisher-key-id"));
      assert.ok(exposed.has("x-package-scan-status"));
      assert.ok(exposed.has("x-package-files-sha256"));
      assert.ok(exposed.has("x-package-published-at"));
    }

    // Extension details (200 + 304) + 404 should all include CORS headers.
    {
      const url = `${baseUrl}/api/extensions/${encodeURIComponent(published.id)}`;
      const first = await fetch(url, { headers: { Origin: origin } });
      assert.equal(first.status, 200);
      assert.equal(first.headers.get("access-control-allow-origin"), "*");
      const etag = first.headers.get("etag");
      assert.ok(etag);

      const cached = await fetch(url, { headers: { Origin: origin, "If-None-Match": etag } });
      assert.equal(cached.status, 304);
      assert.equal(cached.headers.get("access-control-allow-origin"), "*");
      assert.equal(cached.headers.get("etag"), etag);

      const missing = await fetch(`${baseUrl}/api/extensions/does-not-exist`, { headers: { Origin: origin } });
      assert.equal(missing.status, 404);
      assert.equal(missing.headers.get("access-control-allow-origin"), "*");
    }

    // Download (200 + 304)
    {
      const url = `${baseUrl}/api/extensions/${encodeURIComponent(published.id)}/download/${encodeURIComponent(
        published.version,
      )}`;
      const first = await fetch(url, { headers: { Origin: origin } });
      assert.equal(first.status, 200);
      assert.equal(first.headers.get("access-control-allow-origin"), "*");
      assert.ok(first.headers.get("x-package-sha256"));
      const etag = first.headers.get("etag");
      assert.ok(etag);
      await first.arrayBuffer();

      const cached = await fetch(url, { headers: { Origin: origin, "If-None-Match": etag } });
      assert.equal(cached.status, 304);
      assert.equal(cached.headers.get("access-control-allow-origin"), "*");
      assert.equal(cached.headers.get("etag"), etag);
    }

    // Preflight should succeed for public GET endpoints.
    {
      const res = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(published.id)}`, {
        method: "OPTIONS",
        headers: {
          Origin: origin,
          "Access-Control-Request-Method": "GET",
          "Access-Control-Request-Headers": "If-None-Match",
        },
      });
      assert.equal(res.status, 204);
      assert.equal(res.headers.get("access-control-allow-origin"), "*");
      assert.equal(res.headers.get("access-control-allow-methods"), "GET, OPTIONS");
      assert.equal(res.headers.get("access-control-allow-headers"), "If-None-Match");
      assert.equal(res.headers.get("access-control-max-age"), "86400");
    }
  } finally {
    await closeServer(server);
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});
