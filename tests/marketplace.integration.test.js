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
import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";
import extensionPackagePkg from "../shared/extension-package/index.js";
import signingPkg from "../shared/crypto/signing.js";

const { createMarketplaceServer } = marketplaceServerPkg;
const { publishExtension, packageExtension } = publisherPkg;
const { ExtensionHost } = extensionHostPkg;
const { createExtensionPackageV1, createExtensionPackageV2, verifyExtensionPackageV2 } = extensionPackagePkg;
const { signBytes, verifyBytesSignature, sha256 } = signingPkg;

const EXTENSION_TIMEOUT_MS = 20_000;

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

async function patchManifest(extensionDir, patch) {
  const manifestPath = path.join(extensionDir, "package.json");
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  Object.assign(manifest, patch);
  await fs.writeFile(manifestPath, JSON.stringify(manifest, null, 2));
}

function keyIdFromPublicKeyPem(publicKeyPem) {
  const key = crypto.createPublicKey(publicKeyPem);
  const der = key.export({ type: "spki", format: "der" });
  return crypto.createHash("sha256").update(der).digest("hex");
}

test("desktop runtime: auto-load installed extensions + hot reload on update", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-runtime-"));
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

    const marketplaceClient = new MarketplaceClient({ baseUrl });
    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    await manager.install(extensionId);

    const runtime = new ExtensionHostManager({
      extensionsDir,
      statePath,
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
      extensionStoragePath: path.join(tmpRoot, "storage.json"),
      permissionPrompt: async () => true,
      activationTimeoutMs: EXTENSION_TIMEOUT_MS,
      commandTimeoutMs: EXTENSION_TIMEOUT_MS,
      customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
      dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
    });
    await runtime.startup();

    runtime.spreadsheet.setCell(0, 0, 1);
    runtime.spreadsheet.setCell(0, 1, 2);
    runtime.spreadsheet.setCell(1, 0, 3);
    runtime.spreadsheet.setCell(1, 1, 4);
    runtime.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const resultV1 = await runtime.executeCommand("sampleHello.sumSelection");
    assert.equal(resultV1, 10);

    const extSourceV11 = path.join(tmpRoot, "ext-v1.1.0");
    await copyDir(sampleExtensionSrc, extSourceV11);
    await writeManifestVersion(extSourceV11, "1.1.0");

    // Change a contribution and behavior so the test can prove we reloaded both the manifest
    // and the runtime worker after updating the installed files.
    const manifestV11Path = path.join(extSourceV11, "package.json");
    const manifestV11 = JSON.parse(await fs.readFile(manifestV11Path, "utf8"));
    const sumCommand = (manifestV11.contributes?.commands ?? []).find((cmd) => cmd.command === "sampleHello.sumSelection");
    assert.ok(sumCommand);
    sumCommand.title = "Sum Selection v1.1.0";
    await fs.writeFile(manifestV11Path, JSON.stringify(manifestV11, null, 2));

    const distPath = path.join(extSourceV11, "dist", "extension.js");
    const distText = await fs.readFile(distPath, "utf8");
    const commandMarker = 'registerCommand("sampleHello.sumSelection"';
    const commandIdx = distText.indexOf(commandMarker);
    assert.ok(commandIdx >= 0);
    const needle = "return sum;";
    const returnIdx = distText.indexOf(needle, commandIdx);
    assert.ok(returnIdx >= 0);
    const patched =
      distText.slice(0, returnIdx) + "return sum + 1;" + distText.slice(returnIdx + needle.length);
    assert.notEqual(patched, distText);
    await fs.writeFile(distPath, patched);

    await publishExtension({
      extensionDir: extSourceV11,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    // Unload first so the update doesn't rewrite files under a live extension worker.
    await runtime.unloadExtension(extensionId);
    await manager.update(extensionId);
    await runtime.reloadExtension(extensionId);

    const contributions = runtime.listContributions();
    const cmdV11 = contributions.commands.find((cmd) => cmd.command === "sampleHello.sumSelection");
    assert.ok(cmdV11);
    assert.equal(cmdV11.title, "Sum Selection v1.1.0");

    runtime.spreadsheet.setCell(0, 0, 1);
    runtime.spreadsheet.setCell(0, 1, 2);
    runtime.spreadsheet.setCell(1, 0, 3);
    runtime.spreadsheet.setCell(1, 1, 4);
    runtime.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const resultV11 = await runtime.executeCommand("sampleHello.sumSelection");
    assert.equal(resultV11, 11);

    await runtime.dispose();
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

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
    await assert.doesNotReject(() =>
      fs.stat(path.join(dataDir, "packages", extensionId, "1.0.0.fextpkg"))
    );

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
      activationTimeoutMs: EXTENSION_TIMEOUT_MS,
      commandTimeoutMs: EXTENSION_TIMEOUT_MS,
      customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
      dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
    });
 
    const installedPath = path.join(extensionsDir, extensionId);
    await manager.loadIntoHost(host, extensionId);

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
    await assert.doesNotReject(() =>
      fs.stat(path.join(dataDir, "packages", extensionId, "1.1.0.fextpkg"))
    );

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
      activationTimeoutMs: EXTENSION_TIMEOUT_MS,
      commandTimeoutMs: EXTENSION_TIMEOUT_MS,
      customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
      dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
    });
 
    await manager.loadIntoHost(host2, extensionId);
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

test("tampered marketplace download does not clobber an existing install", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-tamper-"));
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

    const marketplaceClient = new MarketplaceClient({ baseUrl });
    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    await manager.install(extensionId, "1.0.0");

    const installedBefore = await manager.getInstalled(extensionId);
    assert.ok(installedBefore);
    assert.equal(installedBefore.version, "1.0.0");

    // Publish an update, then tamper with the stored package bytes on disk to simulate a
    // malicious/bitflipped marketplace download.
    const extSourceV11 = path.join(tmpRoot, "ext-v1.1.0");
    await copyDir(sampleExtensionSrc, extSourceV11);
    await writeManifestVersion(extSourceV11, "1.1.0");
    await publishExtension({
      extensionDir: extSourceV11,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const packagePath = path.join(dataDir, "packages", extensionId, "1.1.0.fextpkg");
    const pkgBytes = await fs.readFile(packagePath);
    const tampered = Buffer.from(pkgBytes);
    tampered[Math.floor(tampered.length / 2)] ^= 0x01;
    await fs.writeFile(packagePath, tampered);

    await assert.rejects(() => manager.update(extensionId), /signature|checksum|sha256 mismatch|tar checksum/i);

    // The failed update should not have removed or partially overwritten the existing install.
    const installedAfter = await manager.getInstalled(extensionId);
    assert.ok(installedAfter);
    assert.equal(installedAfter.version, "1.0.0");

    const installedPackageJson = JSON.parse(
      await fs.readFile(path.join(extensionsDir, extensionId, "package.json"), "utf8"),
    );
    assert.equal(installedPackageJson.version, "1.0.0");
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("installed extension tampering is detected, quarantines execution, and repair reinstalls", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-integrity-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  let runtime = null;
  let runtime2 = null;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPem);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
    const baseManifest = JSON.parse(JSON.stringify(manifest));
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
      extensionDir: extSource,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const marketplaceClient = new MarketplaceClient({ baseUrl });
    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    await manager.install(extensionId);

    const tamperPath = path.join(extensionsDir, extensionId, "dist", "extension.js");
    await fs.appendFile(tamperPath, "\n// tampered\n", "utf8");

    runtime = new ExtensionHostManager({
      extensionsDir,
      statePath,
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
      extensionStoragePath: path.join(tmpRoot, "storage.json"),
      permissionPrompt: async () => true,
      activationTimeoutMs: EXTENSION_TIMEOUT_MS,
      commandTimeoutMs: EXTENSION_TIMEOUT_MS,
      customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
      dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
    });

    // Startup should quarantine the tampered install (mark state corrupted) but allow the
    // runtime to continue booting so other extensions can still run.
    await runtime.startup();

    // The tampered extension should not be loaded into the runtime.
    await assert.rejects(() => runtime.executeCommand("sampleHello.sumSelection"), /Unknown command/i);
    const contributions = runtime.listContributions();
    assert.ok(!contributions.commands.some((cmd) => cmd.command === "sampleHello.sumSelection"));

    const corruptedState = await manager.getInstalled(extensionId);
    assert.ok(corruptedState?.corrupted);
    assert.match(
      String(corruptedState?.corruptedReason ?? ""),
      /Checksum mismatch|Size mismatch|Unexpected file|Missing expected file|integrity/i,
    );

    await manager.repair(extensionId);

    runtime2 = new ExtensionHostManager({
      extensionsDir,
      statePath,
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
      extensionStoragePath: path.join(tmpRoot, "storage.json"),
      permissionPrompt: async () => true,
      activationTimeoutMs: EXTENSION_TIMEOUT_MS,
      commandTimeoutMs: EXTENSION_TIMEOUT_MS,
      customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
      dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
    });

    await runtime2.startup();
    runtime2.spreadsheet.setCell(0, 0, 1);
    runtime2.spreadsheet.setCell(0, 1, 2);
    runtime2.spreadsheet.setCell(1, 0, 3);
    runtime2.spreadsheet.setCell(1, 1, 4);
    runtime2.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const result = await runtime2.executeCommand("sampleHello.sumSelection");
    assert.equal(result, 10);

    const repaired = await manager.getInstalled(extensionId);
    assert.ok(repaired);
    assert.ok(!repaired.corrupted);
  } finally {
    if (runtime) await runtime.dispose().catch(() => {});
    if (runtime2) await runtime2.dispose().catch(() => {});
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publish-bin accepts v2 extension packages without X-Package-Signature", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-v2-nosig-"));
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

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    const packaged = await packageExtension(extSource, { privateKeyPem });
    const pkgSha = sha256(packaged.packageBytes);

    const invalidShaRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
        "X-Package-Sha256": "not-a-sha",
      },
      body: packaged.packageBytes,
    });
    assert.equal(invalidShaRes.status, 400);
    const invalidShaBody = await invalidShaRes.json();
    assert.match(String(invalidShaBody?.error || ""), /invalid x-package-sha256/i);

    const badShaRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
        "X-Package-Sha256": "0".repeat(64),
      },
      body: packaged.packageBytes,
    });
    assert.equal(badShaRes.status, 400);
    const badShaBody = await badShaRes.json();
    assert.match(String(badShaBody?.error || ""), /x-package-sha256/i);

    const publishRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
        "X-Package-Sha256": pkgSha,
      },
      body: packaged.packageBytes,
    });
    assert.equal(publishRes.status, 200);
    assert.equal(publishRes.headers.get("cache-control"), "no-store");
    const published = await publishRes.json();
    assert.deepEqual(published, { id: extensionId, version: manifest.version });
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publish-bin rejects invalid manifests (matches client validation)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-invalid-manifest-"));
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

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
    const baseManifest = JSON.parse(JSON.stringify(manifest));

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

    // Patch the on-disk manifest to be invalid and bypass the extension-publisher's local
    // validation by building v2 packages directly.
    async function publishInvalidManifest(patch, expectedErrorRe) {
      await fs.writeFile(
        path.join(extSource, "package.json"),
        JSON.stringify({ ...baseManifest, ...patch }, null, 2),
      );
      const packageBytes = await createExtensionPackageV2(extSource, { privateKeyPem });

      const publishRes = await fetch(`${baseUrl}/api/publish-bin`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${publisherToken}`,
          "Content-Type": "application/vnd.formula.extension-package",
          "X-Package-Sha256": sha256(packageBytes),
        },
        body: packageBytes,
      });
      assert.equal(publishRes.status, 400);
      const body = await publishRes.json();
      assert.match(String(body?.error || ""), expectedErrorRe);
    }

    await publishInvalidManifest({ permissions: ["totally.not.real"] }, /invalid permission/i);

    await publishInvalidManifest(
      {
        permissions: baseManifest.permissions ?? [],
        activationEvents: ["onView:missing.panel"],
        contributes: { panels: [] },
      },
      /unknown view\/panel/i
    );

    await publishInvalidManifest(
      {
        permissions: baseManifest.permissions ?? [],
        activationEvents: ["onCustomFunction:missing.func"],
        contributes: { customFunctions: [] },
      },
      /unknown custom function/i
    );

    await publishInvalidManifest(
      {
        permissions: baseManifest.permissions ?? [],
        activationEvents: ["onDataConnector:missing.connector"],
        contributes: { dataConnectors: [] },
      },
      /unknown data connector/i
    );

    await publishInvalidManifest(
      {
        permissions: baseManifest.permissions ?? [],
        activationEvents: [],
        contributes: { commands: [] },
        browser: "./dist/missing-browser.mjs",
      },
      /browser entrypoint is missing/i
    );

    await publishInvalidManifest(
      {
        permissions: baseManifest.permissions ?? [],
        activationEvents: [],
        contributes: { commands: [] },
        module: "./dist/missing-module.mjs",
      },
      /module entrypoint is missing/i
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publish-bin accepts v1 extension packages with detached X-Package-Signature", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-v1-bin-"));
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

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    const packageBytes = await createExtensionPackageV1(extSource);
    const signatureBase64 = signBytes(packageBytes, privateKeyPem);
    const pkgSha = sha256(packageBytes);

    const badShaRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
        "X-Package-Signature": signatureBase64,
        "X-Package-Sha256": "0".repeat(64),
      },
      body: packageBytes,
    });
    assert.equal(badShaRes.status, 400);
    const badShaBody = await badShaRes.json();
    assert.match(String(badShaBody?.error || ""), /x-package-sha256/i);

    const publishRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
        "X-Package-Signature": signatureBase64,
        "X-Package-Sha256": pkgSha,
      },
      body: packageBytes,
    });
    assert.equal(publishRes.status, 200);
    assert.equal(publishRes.headers.get("cache-control"), "no-store");
    const published = await publishRes.json();
    assert.deepEqual(published, { id: extensionId, version: manifest.version });

    const missingSigRes = await fetch(`${baseUrl}/api/publish-bin`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/vnd.formula.extension-package",
      },
      body: packageBytes,
    });
    assert.equal(missingSigRes.status, 400);
    const missingError = await missingSigRes.json();
    assert.match(String(missingError?.error || ""), /x-package-signature/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publish accepts v1 extension packages with detached signatureBase64", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-v1-"));
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

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    const packageBytes = await createExtensionPackageV1(extSource);
    const signatureBase64 = signBytes(packageBytes, privateKeyPem);

    const publishRes = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packageBytes.toString("base64"),
        signatureBase64,
      }),
    });
    assert.equal(publishRes.status, 200);
    const published = await publishRes.json();
    assert.deepEqual(published, { id: extensionId, version: manifest.version });

    const downloadRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent(manifest.version)}`
    );
    assert.equal(downloadRes.status, 200);
    assert.equal(downloadRes.headers.get("x-package-format-version"), "1");
    const downloadedSignature = downloadRes.headers.get("x-package-signature");
    assert.equal(downloadedSignature, signatureBase64);

    const downloadedBytes = Buffer.from(await downloadRes.arrayBuffer());
    assert.ok(verifyBytesSignature(downloadedBytes, downloadedSignature, publicKeyPem));
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("desktop client rejects downloads when X-Package-Sha256 mismatches bytes", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-sha-mismatch-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server, store } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    await publishExtension({ extensionDir: extSource, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });

    const bogusSha = "0".repeat(64);
    await store.db.withTransaction((db) => {
      db.run(`UPDATE extension_versions SET sha256 = ? WHERE extension_id = ? AND version = ?`, [
        bogusSha,
        extensionId,
        manifest.version,
      ]);
    });

    const client = new MarketplaceClient({ baseUrl });
    await assert.rejects(() => client.downloadPackage(extensionId, manifest.version), /sha256 mismatch/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    store.close();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("desktop client rejects downloads with invalid X-Package-Sha256 header", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-sha-invalid-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server, store } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    await publishExtension({ extensionDir: extSource, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });

    await store.db.withTransaction((db) => {
      db.run(`UPDATE extension_versions SET sha256 = ? WHERE extension_id = ? AND version = ?`, [
        "not-a-sha",
        extensionId,
        manifest.version,
      ]);
    });

    const client = new MarketplaceClient({ baseUrl });
    await assert.rejects(() => client.downloadPackage(extensionId, manifest.version), /invalid x-package-sha256/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    store.close();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("marketplace responses include ETag and honor If-None-Match (304)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-etag-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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

    await publishExtension({ extensionDir: extSource, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });

    const extUrl = `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}`;
    const extRes = await fetch(extUrl);
    assert.equal(extRes.status, 200);
    assert.equal(extRes.headers.get("cache-control"), "public, max-age=0, must-revalidate");
    const extEtag = extRes.headers.get("etag");
    assert.ok(extEtag);
    await extRes.text();

    const extRes304 = await fetch(extUrl, { headers: { "If-None-Match": extEtag } });
    assert.equal(extRes304.status, 304);
    assert.equal(extRes304.headers.get("cache-control"), "public, max-age=0, must-revalidate");
    assert.equal(extRes304.headers.get("etag"), extEtag);
    assert.equal(await extRes304.text(), "");

    const dlUrl = `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent(manifest.version)}`;
    const dlRes = await fetch(dlUrl);
    assert.equal(dlRes.status, 200);
    assert.equal(dlRes.headers.get("cache-control"), "public, max-age=0, must-revalidate");
    const pkgEtag = dlRes.headers.get("etag");
    const pkgSha = dlRes.headers.get("x-package-sha256");
    assert.ok(pkgEtag);
    assert.ok(pkgSha);
    await dlRes.arrayBuffer();

    const dlRes304 = await fetch(dlUrl, { headers: { "If-None-Match": pkgEtag } });
    assert.equal(dlRes304.status, 304);
    assert.equal(dlRes304.headers.get("cache-control"), "public, max-age=0, must-revalidate");
    assert.equal(dlRes304.headers.get("etag"), pkgEtag);
    assert.equal(dlRes304.headers.get("x-package-sha256"), pkgSha);
    assert.equal(await dlRes304.arrayBuffer().then((b) => b.byteLength), 0);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("extension metadata ETag changes when publisher public key changes", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-etag-key-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const keyA = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemA = keyA.publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPemA = keyA.privateKey.export({ type: "pkcs8", format: "pem" });

    const keyB = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemB = keyB.publicKey.export({ type: "spki", format: "pem" });

    const publisherToken = "publisher-token";
    const privateKeyPath = path.join(tmpRoot, "publisher-private.pem");
    await fs.writeFile(privateKeyPath, privateKeyPemA);

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regA = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemA,
        verified: true,
      }),
    });
    assert.equal(regA.status, 200);

    await publishExtension({
      extensionDir: extSource,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const extUrl = `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}`;
    const first = await fetch(extUrl);
    assert.equal(first.status, 200);
    const etagA = first.headers.get("etag");
    assert.ok(etagA);
    const bodyA = await first.json();
    assert.equal(bodyA.publisherPublicKeyPem, publicKeyPemA);

    const regB = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemB,
        verified: true,
      }),
    });
    assert.equal(regB.status, 200);

    const second = await fetch(extUrl, { headers: { "If-None-Match": etagA } });
    assert.equal(second.status, 200);
    const etagB = second.headers.get("etag");
    assert.ok(etagB && etagB !== etagA);
    const bodyB = await second.json();
    assert.equal(bodyB.publisherPublicKeyPem, publicKeyPemB);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publisher key rotation preserves verification + installs across historical versions", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-key-rotation-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const keyA = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemA = keyA.publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPemA = keyA.privateKey.export({ type: "pkcs8", format: "pem" });
    const keyIdA = keyIdFromPublicKeyPem(publicKeyPemA);

    const keyB = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemB = keyB.publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPemB = keyB.privateKey.export({ type: "pkcs8", format: "pem" });
    const keyIdB = keyIdFromPublicKeyPem(publicKeyPemB);

    const publisherToken = "publisher-token";

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");

    const extV1 = path.join(tmpRoot, "ext-v1");
    await copyDir(sampleExtensionSrc, extV1);
    const manifestV1 = JSON.parse(await fs.readFile(path.join(extV1, "package.json"), "utf8"));
    const extensionId = `${manifestV1.publisher}.${manifestV1.name}`;

    const regA = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifestV1.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemA,
        verified: true,
      }),
    });
    assert.equal(regA.status, 200);

    // Publish a v1 package signed with key A.
    const packageBytesV1 = await createExtensionPackageV1(extV1);
    const signatureBase64V1 = signBytes(packageBytesV1, privateKeyPemA);
    const publishV1 = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packageBytesV1.toString("base64"),
        signatureBase64: signatureBase64V1,
      }),
    });
    assert.equal(publishV1.status, 200);
    assert.deepEqual(await publishV1.json(), { id: extensionId, version: "1.0.0" });

    // Publish a v2 package signed with key A.
    const extV2A = path.join(tmpRoot, "ext-v2-a");
    await copyDir(sampleExtensionSrc, extV2A);
    await writeManifestVersion(extV2A, "1.1.0");
    const packagedA = await packageExtension(extV2A, { privateKeyPem: privateKeyPemA });
    const publishV2A = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packagedA.packageBytes.toString("base64"),
      }),
    });
    assert.equal(publishV2A.status, 200);
    assert.deepEqual(await publishV2A.json(), { id: extensionId, version: "1.1.0" });

    // Rotate publisher primary key to B (keep A in history).
    const regB = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifestV1.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemB,
        verified: true,
      }),
    });
    assert.equal(regB.status, 200);

    // Publish a v2 package signed with key B.
    const extV2B = path.join(tmpRoot, "ext-v2-b");
    await copyDir(sampleExtensionSrc, extV2B);
    await writeManifestVersion(extV2B, "1.2.0");
    const packagedB = await packageExtension(extV2B, { privateKeyPem: privateKeyPemB });
    const publishV2B = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packagedB.packageBytes.toString("base64"),
      }),
    });
    assert.equal(publishV2B.status, 200);
    assert.deepEqual(await publishV2B.json(), { id: extensionId, version: "1.2.0" });

    const marketplaceClient = new MarketplaceClient({ baseUrl });

    const ext = await marketplaceClient.getExtension(extensionId);
    assert.ok(ext);
    assert.equal(String(ext.publisherPublicKeyPem || "").trim(), String(publicKeyPemB).trim());
    assert.ok(Array.isArray(ext.publisherKeys));
    const keysById = new Map(ext.publisherKeys.map((k) => [k.id, k]));
    assert.ok(keysById.has(keyIdA));
    assert.ok(keysById.has(keyIdB));
    assert.equal(keysById.get(keyIdA).revoked, false);
    assert.equal(keysById.get(keyIdB).revoked, false);

    const downloadV1 = await marketplaceClient.downloadPackage(extensionId, "1.0.0");
    assert.ok(downloadV1);
    assert.equal(downloadV1.publisherKeyId, keyIdA);
    assert.ok(verifyBytesSignature(downloadV1.bytes, downloadV1.signatureBase64, publicKeyPemA));
    assert.ok(!verifyBytesSignature(downloadV1.bytes, downloadV1.signatureBase64, publicKeyPemB));

    const downloadV2A = await marketplaceClient.downloadPackage(extensionId, "1.1.0");
    assert.ok(downloadV2A);
    assert.equal(downloadV2A.publisherKeyId, keyIdA);
    assert.doesNotThrow(() => verifyExtensionPackageV2(downloadV2A.bytes, publicKeyPemA));
    assert.throws(() => verifyExtensionPackageV2(downloadV2A.bytes, publicKeyPemB));

    const downloadV2B = await marketplaceClient.downloadPackage(extensionId, "1.2.0");
    assert.ok(downloadV2B);
    assert.equal(downloadV2B.publisherKeyId, keyIdB);
    assert.doesNotThrow(() => verifyExtensionPackageV2(downloadV2B.bytes, publicKeyPemB));
    assert.throws(() => verifyExtensionPackageV2(downloadV2B.bytes, publicKeyPemA));

    // Ensure clients can still install older versions after rotation (primary key != signing key).
    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    const installedV1 = await manager.install(extensionId, "1.0.0");
    assert.equal(installedV1.version, "1.0.0");
    const installedV2A = await manager.install(extensionId, "1.1.0");
    assert.equal(installedV2A.version, "1.1.0");
    const installedV2B = await manager.install(extensionId, "1.2.0");
    assert.equal(installedV2B.version, "1.2.0");
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("revoked signing keys are rejected for publish + install", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-key-revoke-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const keyA = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemA = keyA.publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPemA = keyA.privateKey.export({ type: "pkcs8", format: "pem" });
    const keyIdA = keyIdFromPublicKeyPem(publicKeyPemA);

    const keyB = crypto.generateKeyPairSync("ed25519");
    const publicKeyPemB = keyB.publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPemB = keyB.privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extV2A = path.join(tmpRoot, "ext-v2-a");
    await copyDir(sampleExtensionSrc, extV2A);
    const manifest = JSON.parse(await fs.readFile(path.join(extV2A, "package.json"), "utf8"));
    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const regA = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemA,
        verified: true,
      }),
    });
    assert.equal(regA.status, 200);

    const packagedA = await packageExtension(extV2A, { privateKeyPem: privateKeyPemA });
    const publishA = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packagedA.packageBytes.toString("base64"),
      }),
    });
    assert.equal(publishA.status, 200);

    // Rotate to key B.
    const regB = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifest.publisher,
        token: publisherToken,
        publicKeyPem: publicKeyPemB,
        verified: true,
      }),
    });
    assert.equal(regB.status, 200);

    // Publish a version signed with key B.
    const extV2B = path.join(tmpRoot, "ext-v2-b");
    await copyDir(sampleExtensionSrc, extV2B);
    await writeManifestVersion(extV2B, "1.1.0");
    const packagedB = await packageExtension(extV2B, { privateKeyPem: privateKeyPemB });
    const publishB = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packagedB.packageBytes.toString("base64"),
      }),
    });
    assert.equal(publishB.status, 200);

    const revoke = await fetch(
      `${baseUrl}/api/publishers/${encodeURIComponent(manifest.publisher)}/keys/${encodeURIComponent(keyIdA)}/revoke`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${adminToken}`,
        },
      }
    );
    assert.equal(revoke.status, 200);

    // Publishing with the revoked key should now fail.
    const extV2A2 = path.join(tmpRoot, "ext-v2-a2");
    await copyDir(sampleExtensionSrc, extV2A2);
    await writeManifestVersion(extV2A2, "1.2.0");
    const packagedA2 = await packageExtension(extV2A2, { privateKeyPem: privateKeyPemA });
    const publishA2 = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packagedA2.packageBytes.toString("base64"),
      }),
    });
    assert.equal(publishA2.status, 400);

    const marketplaceClient = new MarketplaceClient({ baseUrl });
    const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
    await assert.rejects(() => manager.install(extensionId, "1.0.0"), /signature verification failed/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("desktop client uses on-disk cache + If-None-Match for package downloads", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-client-cache-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const cacheDir = path.join(tmpRoot, "client-cache");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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
      extensionDir: extSource,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const client = new MarketplaceClient({ baseUrl, cacheDir });
    const first = await client.downloadPackage(extensionId, manifest.version);
    assert.ok(first);

    const second = await client.downloadPackage(extensionId, manifest.version);
    assert.ok(second);
    assert.equal(second.sha256, first.sha256);
    assert.ok(second.bytes.equals(first.bytes));

    const metricsRes = await fetch(`${baseUrl}/api/internal/metrics`);
    assert.equal(metricsRes.status, 200);
    const metricsText = await metricsRes.text();
    assert.match(
      metricsText,
      /marketplace_http_requests_total\{method="GET",route="\/api\/extensions\/:id\/download\/:version",status="200"\} 1/
    );
    assert.match(
      metricsText,
      /marketplace_http_requests_total\{method="GET",route="\/api\/extensions\/:id\/download\/:version",status="304"\} 1/
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("desktop client uses on-disk cache + If-None-Match for extension metadata", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-client-meta-cache-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");
  const cacheDir = path.join(tmpRoot, "client-cache");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { getExtensionPerIpPerMinute: 0 } });

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
    const extSource = path.join(tmpRoot, "ext");
    await copyDir(sampleExtensionSrc, extSource);

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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
      extensionDir: extSource,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const client = new MarketplaceClient({ baseUrl, cacheDir });
    const first = await client.getExtension(extensionId);
    assert.ok(first);

    const second = await client.getExtension(extensionId);
    assert.ok(second);
    assert.equal(second.id, extensionId);

    const metricsRes = await fetch(`${baseUrl}/api/internal/metrics`);
    assert.equal(metricsRes.status, 200);
    const metricsText = await metricsRes.text();
    assert.match(
      metricsText,
      /marketplace_http_requests_total\{method="GET",route="\/api\/extensions\/:id",status="200"\} 1/
    );
    assert.match(
      metricsText,
      /marketplace_http_requests_total\{method="GET",route="\/api\/extensions\/:id",status="304"\} 1/
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("search matches multi-token queries across different fields", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-search-tokens-"));
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
    const extSource = path.join(tmpRoot, "ext-search");
    await copyDir(sampleExtensionSrc, extSource);

    await patchManifest(extSource, {
      name: "sample-linter",
      displayName: "Sample Linter",
      description: "A linter extension",
      tags: ["python"],
    });

    const manifest = JSON.parse(await fs.readFile(path.join(extSource, "package.json"), "utf8"));
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
      extensionDir: extSource,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const client = new MarketplaceClient({ baseUrl });
    const result = await client.search({ q: "python linter" });
    assert.ok(result.results.some((r) => r.id === extensionId));
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("search supports category/tag/verified/featured filters", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-search-filters-"));
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
    const extA = path.join(tmpRoot, "ext-a");
    const extB = path.join(tmpRoot, "ext-b");
    await copyDir(sampleExtensionSrc, extA);
    await copyDir(sampleExtensionSrc, extB);

    await patchManifest(extA, { name: "sample-filter-a", categories: ["Utilities"], tags: ["foo"] });
    await patchManifest(extB, { name: "sample-filter-b", categories: ["Themes"], tags: ["bar"] });

    const manifestA = JSON.parse(await fs.readFile(path.join(extA, "package.json"), "utf8"));
    const manifestB = JSON.parse(await fs.readFile(path.join(extB, "package.json"), "utf8"));
    const idA = `${manifestA.publisher}.${manifestA.name}`;
    const idB = `${manifestB.publisher}.${manifestB.name}`;

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: manifestA.publisher,
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    await publishExtension({ extensionDir: extA, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });
    await publishExtension({ extensionDir: extB, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });

    const featureRes = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(idB)}/flags`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ featured: true }),
    });
    assert.equal(featureRes.status, 200);

    const client = new MarketplaceClient({ baseUrl });

    const byCategory = await client.search({ category: "utilities" });
    assert.deepEqual(byCategory.results.map((r) => r.id).sort(), [idA]);

    const byTag = await client.search({ tag: "bar" });
    assert.deepEqual(byTag.results.map((r) => r.id).sort(), [idB]);

    const byFeatured = await client.search({ featured: true });
    assert.deepEqual(byFeatured.results.map((r) => r.id).sort(), [idB]);

    const byVerified = await client.search({ verified: true });
    assert.deepEqual(new Set(byVerified.results.map((r) => r.id)), new Set([idA, idB]));
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

    const downloadRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.0.0")}`
    );
    assert.equal(downloadRes.status, 404);

    const auditRes = await fetch(`${baseUrl}/api/admin/audit?limit=100`, {
      headers: { Authorization: `Bearer ${adminToken}` },
    });
    assert.equal(auditRes.status, 200);
    const audit = await auditRes.json();
    assert.ok(audit.entries.some((e) => e.action === "extension.flags"));
    assert.ok(audit.entries.some((e) => e.action === "extension.version.flags"));
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("moderation: yanking latest version falls back to previous for latestVersion", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-yank-latest-"));
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
    const extSourceV11 = path.join(tmpRoot, "ext-v1.1.0");
    await copyDir(sampleExtensionSrc, extSourceV1);
    await copyDir(sampleExtensionSrc, extSourceV11);
    await writeManifestVersion(extSourceV11, "1.1.0");

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
    await publishExtension({
      extensionDir: extSourceV11,
      marketplaceUrl: baseUrl,
      token: publisherToken,
      privateKeyPemOrPath: privateKeyPath,
    });

    const yankRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/versions/${encodeURIComponent("1.1.0")}/flags`,
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

    const client = new MarketplaceClient({ baseUrl });
    const ext = await client.getExtension(extensionId);
    assert.ok(ext);
    assert.equal(ext.latestVersion, "1.0.0");
    assert.ok(ext.versions.some((v) => v.version === "1.1.0" && v.yanked));

    const search = await client.search({ q: "sample" });
    const hit = search.results.find((r) => r.id === extensionId);
    assert.ok(hit);
    assert.equal(hit.latestVersion, "1.0.0");

    const downloadYanked = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.1.0")}`
    );
    assert.equal(downloadYanked.status, 404);
    const downloadOk = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.0.0")}`
    );
    assert.equal(downloadOk.status, 200);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("publish rejects duplicate version under concurrent publish attempts", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-dup-publish-"));
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

    const results = await Promise.allSettled([
      publishExtension({
        extensionDir: extSourceV1,
        marketplaceUrl: baseUrl,
        token: publisherToken,
        privateKeyPemOrPath: privateKeyPath,
      }),
      publishExtension({
        extensionDir: extSourceV1,
        marketplaceUrl: baseUrl,
        token: publisherToken,
        privateKeyPemOrPath: privateKeyPath,
      }),
    ]);

    assert.equal(results.filter((r) => r.status === "fulfilled").length, 1);
    assert.equal(results.filter((r) => r.status === "rejected").length, 1);
    const rejection = results.find((r) => r.status === "rejected");
    assert.ok(rejection && rejection.status === "rejected");
    assert.match(String(rejection.reason?.message || rejection.reason), /409|already published/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("moderation: blocked extensions are hidden from getExtension/search/download", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-blocked-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    assert.ok(await client.getExtension(extensionId));

    const blockRes = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/flags`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ blocked: true }),
    });
    assert.equal(blockRes.status, 200);

    const after = await client.search({ q: "sample" });
    assert.ok(!after.results.some((r) => r.id === extensionId));
    assert.equal(await client.getExtension(extensionId), null);

    const downloadRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.0.0")}`,
      { headers: { "X-Forwarded-For": "203.0.113.11" } }
    );
    assert.equal(downloadRes.status, 404);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("moderation: malicious extensions are hidden from getExtension/search/download", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-malicious-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({ dataDir, adminToken, rateLimits: { downloadPerIpPerMinute: 0 } });

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
    assert.ok(await client.getExtension(extensionId));

    const maliciousRes = await fetch(`${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/flags`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ malicious: true }),
    });
    assert.equal(maliciousRes.status, 200);

    const after = await client.search({ q: "sample" });
    assert.ok(!after.results.some((r) => r.id === extensionId));
    assert.equal(await client.getExtension(extensionId), null);

    const downloadRes = await fetch(
      `${baseUrl}/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent("1.0.0")}`,
      { headers: { "X-Forwarded-For": "203.0.113.12" } }
    );
    assert.equal(downloadRes.status, 404);
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
    const retryAfter = Number(limited.headers.get("retry-after") || "0");
    assert.ok(Number.isFinite(retryAfter) && retryAfter > 0);
    // 5 requests/minute -> 1 token every ~12s. Keep a wide bound to avoid flakes.
    assert.ok(retryAfter <= 20);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("rate limiting: /api/publish enforces per-publisher token limits", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-publish-ratelimit-"));
  const dataDir = path.join(tmpRoot, "marketplace-data");

  const adminToken = "admin-secret";
  const { server } = await createMarketplaceServer({
    dataDir,
    adminToken,
    rateLimits: { publishPerPublisherPerMinute: 2 },
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
    const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

    const publisherToken = "publisher-token";

    const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
    const extA = path.join(tmpRoot, "ext-a");
    const extB = path.join(tmpRoot, "ext-b");
    const extC = path.join(tmpRoot, "ext-c");
    await copyDir(sampleExtensionSrc, extA);
    await copyDir(sampleExtensionSrc, extB);
    await copyDir(sampleExtensionSrc, extC);
    await writeManifestVersion(extB, "1.1.0");
    await writeManifestVersion(extC, "1.2.0");

    const manifest = JSON.parse(await fs.readFile(path.join(extA, "package.json"), "utf8"));

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

    const pkgA = await packageExtension(extA, { privateKeyPem });
    const pkgB = await packageExtension(extB, { privateKeyPem });
    const pkgC = await packageExtension(extC, { privateKeyPem });

    for (const pkg of [pkgA, pkgB]) {
      const res = await fetch(`${baseUrl}/api/publish`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${publisherToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ packageBase64: pkg.packageBytes.toString("base64") }),
      });
      assert.equal(res.status, 200);
    }

    const limited = await fetch(`${baseUrl}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${publisherToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ packageBase64: pkgC.packageBytes.toString("base64") }),
    });
    assert.equal(limited.status, 429);
    const retryAfter = Number(limited.headers.get("retry-after") || "0");
    assert.ok(Number.isFinite(retryAfter) && retryAfter > 0);
    // 2 requests/minute -> 1 token every ~30s. Keep a wide bound to avoid flakes.
    assert.ok(retryAfter <= 35);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("search cursor pagination returns stable pages", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-cursor-"));
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
    const extA = path.join(tmpRoot, "ext-a");
    const extB = path.join(tmpRoot, "ext-b");
    const extC = path.join(tmpRoot, "ext-c");
    await copyDir(sampleExtensionSrc, extA);
    await copyDir(sampleExtensionSrc, extB);
    await copyDir(sampleExtensionSrc, extC);
    await patchManifest(extA, { name: "sample-hello-a" });
    await patchManifest(extB, { name: "sample-hello-b" });
    await patchManifest(extC, { name: "sample-hello-c" });

    const regRes = await fetch(`${baseUrl}/api/publishers/register`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${adminToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        publisher: "formula",
        token: publisherToken,
        publicKeyPem,
        verified: true,
      }),
    });
    assert.equal(regRes.status, 200);

    await publishExtension({ extensionDir: extA, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });
    await new Promise((r) => setTimeout(r, 5));
    await publishExtension({ extensionDir: extB, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });
    await new Promise((r) => setTimeout(r, 5));
    await publishExtension({ extensionDir: extC, marketplaceUrl: baseUrl, token: publisherToken, privateKeyPemOrPath: privateKeyPath });

    const client = new MarketplaceClient({ baseUrl });
    const page1 = await client.search({ q: "", sort: "updated", limit: 2 });
    assert.equal(page1.results.length, 2);
    assert.ok(page1.nextCursor);

    const page2 = await client.search({ q: "", sort: "updated", limit: 2, cursor: page1.nextCursor });
    assert.equal(page2.results.length, 1);

    const ids = [...page1.results, ...page2.results].map((r) => r.id);
    assert.equal(new Set(ids).size, 3);
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
