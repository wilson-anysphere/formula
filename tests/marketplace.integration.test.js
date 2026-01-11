import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

import marketplaceServerPkg from "../services/marketplace/src/server.js";
import publisherPkg from "../tools/extension-publisher/src/publisher.js";
import extensionHostPkg from "../packages/extension-host/src/index.js";
import { MarketplaceClient } from "../apps/desktop/src/marketplace/client.js";
import { ExtensionManager } from "../apps/desktop/src/marketplace/extensionManager.js";

const { createMarketplaceServer } = marketplaceServerPkg;
const { publishExtension } = publisherPkg;
const { ExtensionHost } = extensionHostPkg;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

async function copyDir(srcDir, destDir) {
  await fs.mkdir(destDir, { recursive: true });
  const entries = await fs.readdir(srcDir, { withFileTypes: true });
  for (const entry of entries) {
    const src = path.join(srcDir, entry.name);
    const dest = path.join(destDir, entry.name);
    if (entry.isDirectory()) {
      await copyDir(src, dest);
      continue;
    }
    if (entry.isFile()) {
      await fs.copyFile(src, dest);
    }
  }
}

async function writeManifestVersion(extensionDir, version) {
  const manifestPath = path.join(extensionDir, "package.json");
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  manifest.version = version;
  await fs.writeFile(manifestPath, JSON.stringify(manifest, null, 2));
}

test("marketplace publish → discover → install → verify signature → run command → update", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSourceV1 = path.join(tmpRoot, "ext-v1");
    await copyDir(sampleExtensionSrc, extSourceV1);

    const manifest = JSON.parse(await fs.readFile(path.join(extSourceV1, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    const publishV1 = await publishExtension({
      extensionDir: extSourceV1,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });
    assert.equal(publishV1.id, extensionId);
    assert.equal(publishV1.version, "1.0.0");

    const marketplaceClient = new MarketplaceClient({ baseUrl });
    const search = await marketplaceClient.search({ q: "sample" });
    assert.ok(search.results.some((r) => r.id === extensionId));

    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    await manager.install(extensionId);

    const host = new ExtensionHost({
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
      extensionStoragePath: path.join(tmpRoot, "storage.json"),
      permissionPrompt: async () => true,
    });

    const installedPath = path.join(extensionsDir, extensionId);
    await host.loadExtension(installedPath);

    host.spreadsheet.setCell(0, 0, 1);
    host.spreadsheet.setCell(0, 1, 2);
    host.spreadsheet.setCell(1, 0, 3);
    host.spreadsheet.setCell(1, 1, 4);
    host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const resultV1 = await host.executeCommand("sampleHello.sumSelection");
    assert.equal(resultV1, 10);

    const extSourceV11 = path.join(tmpRoot, "ext-v1.1.0");
    await copyDir(sampleExtensionSrc, extSourceV11);
    await writeManifestVersion(extSourceV11, "1.1.0");

    const publishV11 = await publishExtension({
      extensionDir: extSourceV11,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });
    assert.equal(publishV11.version, "1.1.0");

    const updates = await manager.checkForUpdates();
    assert.deepEqual(updates, [
      { id: extensionId, currentVersion: "1.0.0", latestVersion: "1.1.0" },
    ]);

    await host.dispose();
    await manager.update(extensionId);

    const host2 = new ExtensionHost({
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
      extensionStoragePath: path.join(tmpRoot, "storage.json"),
      permissionPrompt: async () => true,
    });

    await host2.loadExtension(installedPath);
    host2.spreadsheet.setCell(0, 0, 1);
    host2.spreadsheet.setCell(0, 1, 2);
    host2.spreadsheet.setCell(1, 0, 3);
    host2.spreadsheet.setCell(1, 1, 4);
    host2.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const resultV11 = await host2.executeCommand("sampleHello.sumSelection");
    assert.equal(resultV11, 10);

    await host2.dispose();
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("client refuses install when signature verification fails", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-badsig-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSourceV1 = path.join(tmpRoot, "ext-v1");
    await copyDir(sampleExtensionSrc, extSourceV1);
    const manifest = JSON.parse(await fs.readFile(path.join(extSourceV1, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    await publishExtension({
      extensionDir: extSourceV1,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const tamperingClient = new MarketplaceClient({ baseUrl });
    const originalDownload = tamperingClient.downloadPackage.bind(tamperingClient);
    tamperingClient.downloadPackage = async (id, version) => {
      const pkg = await originalDownload(id, version);
      const bytes = Buffer.from(pkg.bytes);
      bytes[0] ^= 0xff;
      return { ...pkg, bytes };
    };

    const manager = new ExtensionManager({ marketplaceClient: tamperingClient, extensionsDir, statePath });
    await assert.rejects(() => manager.install(extensionId), /signature verification failed/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("marketplace persists across restarts and supports concurrent publishes", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-concurrent-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server, store } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSourceV1 = path.join(tmpRoot, "ext-v1");
    const extSourceV11 = path.join(tmpRoot, "ext-v1.1.0");
    const extSourceV12 = path.join(tmpRoot, "ext-v1.2.0");
    await copyDir(sampleExtensionSrc, extSourceV1);
    await copyDir(sampleExtensionSrc, extSourceV11);
    await copyDir(sampleExtensionSrc, extSourceV12);
    await writeManifestVersion(extSourceV11, "1.1.0");
    await writeManifestVersion(extSourceV12, "1.2.0");

    const manifest = JSON.parse(await fs.readFile(path.join(extSourceV1, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    const results = await Promise.all([
      publishExtension({ extensionDir: extSourceV1, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath }),
      publishExtension({ extensionDir: extSourceV11, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath }),
      publishExtension({ extensionDir: extSourceV12, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath }),
    ]);
    assert.deepEqual(
      results.map((r) => r.version).sort(),
      ["1.0.0", "1.1.0", "1.2.0"]
    );

    const client = new MarketplaceClient({ baseUrl });
    const ext = await client.getExtension(extensionId);
    assert.ok(ext);
    assert.equal(ext.latestVersion, "1.2.0");
    assert.equal(ext.versions.length, 3);

    // Restart the server to ensure DB persistence.
    await new Promise((resolve) => server.close(resolve));
    store.close();

    const restarted = await createMarketplaceServer({ dataDir, adminToken });
    await new Promise((resolve) => restarted.server.listen(0, "127.0.0.1", resolve));
    const port2 = restarted.server.address().port;
    const baseUrl2 = `http://127.0.0.1:${port2}`;

    const client2 = new MarketplaceClient({ baseUrl: baseUrl2 });
    const search = await client2.search({ q: "sample" });
    assert.ok(search.results.some((r) => r.id === extensionId));

    await new Promise((resolve) => restarted.server.close(resolve));
    restarted.store.close();
  } finally {
    try {
      await new Promise((resolve) => server.close(resolve));
    } catch {
      // ignore
    }
    store.close();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("moderation: deprecated + yanked extensions are hidden from search", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-moderation-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSourceV1 = path.join(tmpRoot, "ext-v1");
    await copyDir(sampleExtensionSrc, extSourceV1);
    const manifest = JSON.parse(await fs.readFile(path.join(extSourceV1, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    await publishExtension({
      extensionDir: extSourceV1,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const client = new MarketplaceClient({ baseUrl });
    const initial = await client.search({ q: "sample" });
    assert.ok(initial.results.some((r) => r.id === extensionId));

    const deprecateRes = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/flags`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ deprecated: true }),
    });
    assert.equal(deprecateRes.status, 200);

    const afterDeprecate = await client.search({ q: "sample" });
    assert.ok(!afterDeprecate.results.some((r) => r.id === extensionId));

    // Un-deprecate, then yank the only version.
    const undeprecateRes = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/flags`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ deprecated: false }),
    });
    assert.equal(undeprecateRes.status, 200);

    const afterUndeprecate = await client.search({ q: "sample" });
    assert.ok(afterUndeprecate.results.some((r) => r.id === extensionId));

    const yankRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/versions/${encodeURIComponent("1.0.0")}/flags`,
      {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${adminToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ yanked: true }),
      }
    );
    assert.equal(yankRes.status, 200);

    const afterYank = await client.search({ q: "sample" });
    assert.ok(!afterYank.results.some((r) => r.id === extensionId));
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("rate limiting: /api/search enforces per-IP limits", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-ratelimit-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({
    dataDir,
    adminToken,
    rateLimits: { searchPerIpPerMinute: 5 },
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    for (let i = 0; i < 5; i++) {
      const res = await fetch(`${baseUrl}/api/search?q=test`, {
        headers: { "X-Forwarded-For": "203.0.113.10" },
      });
      assert.equal(res.status, 200);
    }

    const limited = await fetch(`${baseUrl}/api/search?q=test`, {
      headers: { "X-Forwarded-For": "203.0.113.10" },
    });
    assert.equal(limited.status, 429);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("download count increments atomically under concurrent downloads", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-downloads-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({
    dataDir,
    adminToken,
    rateLimits: { downloadPerIpPerMinute: 0 },
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSourceV1 = path.join(tmpRoot, "ext-v1");
    await copyDir(sampleExtensionSrc, extSourceV1);
    const manifest = JSON.parse(await fs.readFile(path.join(extSourceV1, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    await publishExtension({
      extensionDir: extSourceV1,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const client = new MarketplaceClient({ baseUrl });
    const before = await client.getExtension(extensionId);
    assert.ok(before);
    assert.equal(before.downloadCount, 0);

    const downloads = 25;
    await Promise.all(
      Array.from({ length: downloads }, async (_v, idx) => {
        const res = await fetch(
          `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.0.0")}`,
          { headers: { "X-Forwarded-For": `198.51.100.${idx}` } }
        );
        assert.equal(res.status, 200);
        await res.arrayBuffer();
      })
    );

    const after = await client.getExtension(extensionId);
    assert.ok(after);
    assert.equal(after.downloadCount, downloads);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});
