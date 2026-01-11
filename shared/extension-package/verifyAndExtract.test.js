const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");

const {
  createExtensionPackageV1,
  createExtensionPackageV2,
  verifyAndExtractExtensionPackage,
} = require("./index");
const { generateEd25519KeyPair, signBytes } = require("../crypto/signing");

const UNIQUE_MARKER = "UNIQUE_VERIFY_AND_EXTRACT_MARKER_7f0b4b5e";

async function createTempExtensionDir({ version }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-verifyextract-"));
  await fs.mkdir(path.join(dir, "dist"), { recursive: true });

  const manifest = {
    name: "temp-ext",
    publisher: "temp-pub",
    version,
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(dir, "dist", "extension.js"), `module.exports = {};\n// ${UNIQUE_MARKER}\n`);
  await fs.writeFile(path.join(dir, "README.md"), "hello\n");

  return { dir, manifest };
}

async function listDirSafe(dir) {
  try {
    return await fs.readdir(dir);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return [];
    throw error;
  }
}

test("verifyAndExtractExtensionPackage extracts verified v2 packages atomically", async (t) => {
  const { dir, manifest } = await createTempExtensionDir({ version: "1.0.0" });
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-install-"));
  const installDir = path.join(tmpRoot, "installed");
  const keys = generateEd25519KeyPair();

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  });

  const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });

  const result = await verifyAndExtractExtensionPackage(pkgBytes, installDir, {
    publicKeyPem: keys.publicKeyPem,
    expectedId: `${manifest.publisher}.${manifest.name}`,
    expectedVersion: manifest.version,
  });

  assert.equal(result.formatVersion, 2);
  assert.deepEqual(result.manifest, manifest);
  assert.ok(typeof result.signatureBase64 === "string" && result.signatureBase64.length > 0);

  const installedJs = await fs.readFile(path.join(installDir, "dist", "extension.js"), "utf8");
  assert.match(installedJs, new RegExp(UNIQUE_MARKER));

  const siblings = await listDirSafe(tmpRoot);
  assert.ok(!siblings.some((name) => name.startsWith(".installed.staging-")));
  assert.ok(!siblings.some((name) => name.startsWith(".installed.backup-")));
});

test("verifyAndExtractExtensionPackage rejects tampered v2 packages and does not create dest dir", async (t) => {
  const { dir, manifest } = await createTempExtensionDir({ version: "1.0.0" });
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-install-tampered-"));
  const installDir = path.join(tmpRoot, "installed");
  const keys = generateEd25519KeyPair();

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  });

  const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
  const tampered = Buffer.from(pkgBytes);
  tampered[Math.floor(tampered.length / 2)] ^= 0x01;

  await assert.rejects(
    () =>
      verifyAndExtractExtensionPackage(tampered, installDir, {
        publicKeyPem: keys.publicKeyPem,
        expectedId: `${manifest.publisher}.${manifest.name}`,
        expectedVersion: manifest.version,
      }),
    /verification failed|checksum|tar checksum|signature/i,
  );

  await assert.rejects(() => fs.stat(installDir), /ENOENT|ENOTDIR/);

  const siblings = await listDirSafe(tmpRoot);
  assert.ok(!siblings.some((name) => name.startsWith(".installed.staging-")));
  assert.ok(!siblings.some((name) => name.startsWith(".installed.backup-")));
});

test("verifyAndExtractExtensionPackage extracts verified v1 packages with detached signature", async (t) => {
  const { dir, manifest } = await createTempExtensionDir({ version: "1.0.0" });
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-install-v1-"));
  const installDir = path.join(tmpRoot, "installed");
  const keys = generateEd25519KeyPair();

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  });

  const pkgBytes = await createExtensionPackageV1(dir);
  const signatureBase64 = signBytes(pkgBytes, keys.privateKeyPem);

  const result = await verifyAndExtractExtensionPackage(pkgBytes, installDir, {
    publicKeyPem: keys.publicKeyPem,
    signatureBase64,
    expectedId: `${manifest.publisher}.${manifest.name}`,
    expectedVersion: manifest.version,
  });

  assert.equal(result.formatVersion, 1);
  assert.deepEqual(result.manifest, manifest);
  assert.equal(result.signatureBase64, signatureBase64);

  const installedPkgJson = JSON.parse(await fs.readFile(path.join(installDir, "package.json"), "utf8"));
  assert.equal(installedPkgJson.version, "1.0.0");
});

