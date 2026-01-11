const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const { createRequire } = require("node:module");
const os = require("node:os");
const path = require("node:path");

const { createExtensionPackageV2, verifyExtensionPackageV2 } = require("../../../../shared/extension-package");
const { generateEd25519KeyPair } = require("../../../../shared/crypto/signing");
const { MarketplaceStore } = require("../store");

const requireFromHere = createRequire(__filename);

async function createTempExtensionDir({ publisher, name, version, jsSource }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-sec-"));
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

async function updateManifestVersion(dir, nextVersion) {
  const packageJsonPath = path.join(dir, "package.json");
  const pkg = JSON.parse(await fs.readFile(packageJsonPath, "utf8"));
  pkg.version = nextVersion;
  await fs.writeFile(packageJsonPath, JSON.stringify(pkg, null, 2));
}

async function updateJsSource(dir, jsSource) {
  await fs.writeFile(path.join(dir, "dist", "extension.js"), jsSource, "utf8");
}

test("publishing triggers a package scan record", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-scan-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "temp-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    const scan = await store.getPackageScan(published.id, published.version);
    assert.ok(scan, "scan record should exist");
    assert.equal(scan.status, "passed");
    assert.ok(scan.scannedAt);

    const auditEntries = await store.listAuditLog({ limit: 20 });
    const publishAudit = auditEntries.find((entry) => entry.action === "extension.publish");
    assert.ok(publishAudit, "publish should be recorded in audit log");
    assert.equal(publishAudit.actor, publisher);
    assert.equal(publishAudit.extensionId, published.id);
    assert.equal(publishAudit.version, published.version);

    const scans = await store.listPackageScans({ status: "passed", limit: 50 });
    assert.ok(scans.some((s) => s.extensionId === published.id && s.version === published.version));
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("revoking a publisher blocks publishing and hides extensions/downloads", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-publisher-revoke-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "revoked-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    await store.revokePublisher(publisher, { revoked: true });

    const hidden = await store.getExtension(published.id);
    assert.equal(hidden, null);

    const adminView = await store.getExtension(published.id, { includeHidden: true });
    assert.ok(adminView);
    assert.equal(adminView.publisherRevoked, true);

    const search = await store.search({ q: published.id });
    assert.equal(search.results.length, 0);

    const downloaded = await store.getPackage(published.id, published.version);
    assert.equal(downloaded, null);

    await assert.rejects(
      () => store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null }),
      /publisher revoked/i
    );
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("scan failed versions are not advertised as latestVersion", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-scan-latest-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "latest-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgV100 = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published100 = await store.publishExtension({ publisher, packageBytes: pkgV100, signatureBase64: null });
    assert.equal(published100.version, "1.0.0");

    await updateManifestVersion(dir, "1.0.1");
    await updateJsSource(
      dir,
      'const cp = require("child_process");\nmodule.exports = { activate() { return cp; } };\n'
    );
    const pkgV101 = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    await store.publishExtension({ publisher, packageBytes: pkgV101, signatureBase64: null });

    const ext = await store.getExtension(published100.id);
    assert.ok(ext);
    assert.equal(ext.latestVersion, "1.0.0");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("malicious packages are flagged and cannot be downloaded", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-malicious-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "malicious-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: 'const cp = require("child_process");\nmodule.exports = { activate() { return cp; } };\n',
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    const scan = await store.getPackageScan(published.id, published.version);
    assert.ok(scan, "scan record should exist");
    assert.equal(scan.status, "failed");
    assert.ok(
      Array.isArray(scan.findings?.findings) && scan.findings.findings.some((f) => f.id === "js.child_process"),
      "scan findings should include child_process usage"
    );

    const downloaded = await store.getPackage(published.id, published.version);
    assert.equal(downloaded, null);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("scan allowlist can permit otherwise-failing packages", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-scan-allowlist-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "allowlisted-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: 'const cp = require("child_process");\nmodule.exports = { activate() { return cp; } };\n',
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir, scanAllowlist: ["JS.CHILD_PROCESS"] });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    const scan = await store.getPackageScan(published.id, published.version);
    assert.ok(scan, "scan record should exist");
    assert.equal(scan.status, "passed");
    assert.ok(!scan.findings?.findings?.some((f) => f.id === "js.child_process"));

    const downloaded = await store.getPackage(published.id, published.version);
    assert.ok(downloaded);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("requireScanPassedForDownload blocks pending scans from metadata/search/download", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-require-scan-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "require-scan-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  try {
    const keys = generateEd25519KeyPair();
    const store = new MarketplaceStore({ dataDir, requireScanPassedForDownload: true });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
    const published = await store.publishExtension({ publisher, packageBytes: pkgBytes, signatureBase64: null });

    // Force the scan status back to pending to simulate an unscanned legacy version.
    await store.db.withTransaction((db) => {
      db.run(`UPDATE package_scans SET status = 'pending', scanned_at = NULL WHERE extension_id = ? AND version = ?`, [
        published.id,
        published.version,
      ]);
    });

    const scan = await store.getPackageScan(published.id, published.version);
    assert.ok(scan);
    assert.equal(scan.status, "pending");

    assert.equal(await store.getExtension(published.id), null);
    const search = await store.search({ q: published.id });
    assert.equal(search.results.length, 0);
    assert.equal(await store.getPackage(published.id, published.version), null);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("publisher key rotation preserves ability to verify old versions", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-key-rotation-"));
  const dataDir = path.join(tmpRoot, "data");
  const publisher = "temp-pub";
  const name = "rotate-ext";

  const { dir } = await createTempExtensionDir({
    publisher,
    name,
    version: "1.0.0",
    jsSource: "module.exports = { activate() {} };\n",
  });

  try {
    const keyA = generateEd25519KeyPair();
    const keyB = generateEd25519KeyPair();
    const keyC = generateEd25519KeyPair();

    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keyA.publicKeyPem,
      verified: true,
    });

    const pkgV100 = await createExtensionPackageV2(dir, { privateKeyPem: keyA.privateKeyPem });
    const published100 = await store.publishExtension({ publisher, packageBytes: pkgV100, signatureBase64: null });

    await store.rotatePublisherPublicKey(publisher, { publicKeyPem: keyB.publicKeyPem, overlapMs: 60 * 60 * 1000 });

    await updateManifestVersion(dir, "1.0.1");
    const pkgV101 = await createExtensionPackageV2(dir, { privateKeyPem: keyB.privateKeyPem });
    const published101 = await store.publishExtension({ publisher, packageBytes: pkgV101, signatureBase64: null });

    await updateManifestVersion(dir, "1.0.2");
    const pkgV102 = await createExtensionPackageV2(dir, { privateKeyPem: keyA.privateKeyPem });
    await store.publishExtension({ publisher, packageBytes: pkgV102, signatureBase64: null });

    await updateManifestVersion(dir, "1.0.3");
    const pkgV103 = await createExtensionPackageV2(dir, { privateKeyPem: keyC.privateKeyPem });
    await assert.rejects(
      () => store.publishExtension({ publisher, packageBytes: pkgV103, signatureBase64: null }),
      /signature verification failed/i
    );

    const ext = await store.getExtension(published100.id);
    assert.ok(ext);
    assert.ok(Array.isArray(ext.publisherKeys));
    const keyPems = ext.publisherKeys.map((k) => String(k.publicKeyPem || "").trim());
    assert.ok(keyPems.includes(keyA.publicKeyPem.trim()));
    assert.ok(keyPems.includes(keyB.publicKeyPem.trim()));

    const downloaded100 = await store.getPackage(published100.id, published100.version);
    assert.ok(downloaded100);
    assert.doesNotThrow(() => verifyExtensionPackageV2(downloaded100.bytes, keyA.publicKeyPem));

    const downloaded101 = await store.getPackage(published101.id, published101.version);
    assert.ok(downloaded101);
    assert.doesNotThrow(() => verifyExtensionPackageV2(downloaded101.bytes, keyB.publicKeyPem));
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});
