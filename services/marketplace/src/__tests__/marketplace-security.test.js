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
