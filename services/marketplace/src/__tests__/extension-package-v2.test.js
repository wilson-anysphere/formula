const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const { createRequire } = require("node:module");
const os = require("node:os");
const path = require("node:path");

const {
  createExtensionPackageV1,
  createExtensionPackageV2,
  readExtensionPackageV2,
  verifyExtensionPackageV2,
} = require("../../../../shared/extension-package");
const { generateEd25519KeyPair, signBytes } = require("../../../../shared/crypto/signing");
const { MarketplaceStore } = require("../store");

const UNIQUE_MARKER = "UNIQUE_PAYLOAD_MARKER_1b2e6b3a";
const requireFromHere = createRequire(__filename);

async function createTempExtensionDir() {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-"));
  await fs.mkdir(path.join(dir, "dist"), { recursive: true });

  const manifest = {
    name: "temp-ext",
    publisher: "temp-pub",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(
    path.join(dir, "dist", "extension.js"),
    `module.exports = { activate() {} };\n// ${UNIQUE_MARKER}\n`,
  );
  await fs.writeFile(path.join(dir, "README.md"), "hello\n");

  return { dir, manifest };
}

function parseTarOctal(buf, offset, length) {
  const str = buf.subarray(offset, offset + length).toString("ascii").replace(/\0.*$/, "").trim();
  if (!str) return 0;
  return Number.parseInt(str, 8);
}

function tarEntryName(header) {
  const name = header.subarray(0, 100).toString("utf8").replace(/\0.*$/, "");
  const prefix = header.subarray(345, 500).toString("utf8").replace(/\0.*$/, "");
  return prefix ? `${prefix}/${name}` : name;
}

function updateTarChecksum(header) {
  header.fill(0x20, 148, 156);
  let sum = 0;
  for (const b of header) sum += b;
  const checksumStr = sum.toString(8).padStart(6, "0");
  header.fill(0, 148, 156);
  header.write(checksumStr, 148, 6, "ascii");
  header[154] = 0;
  header[155] = 0x20;
}

function findTarHeaderOffset(archive, predicate) {
  const BLOCK = 512;
  let offset = 0;
  while (offset + BLOCK <= archive.length) {
    const header = archive.subarray(offset, offset + BLOCK);
    if (header.every((b) => b === 0)) return null;
    const size = parseTarOctal(header, 124, 12);
    const name = tarEntryName(header);
    if (predicate({ name, headerOffset: offset, size })) return offset;
    offset = offset + BLOCK + Math.ceil(size / BLOCK) * BLOCK;
  }
  return null;
}

test("v2 packaging is deterministic (same input → same bytes)", async () => {
  const { dir } = await createTempExtensionDir();
  const { privateKeyPem } = generateEd25519KeyPair();
  try {
    const a = await createExtensionPackageV2(dir, { privateKeyPem });
    const b = await createExtensionPackageV2(dir, { privateKeyPem });
    assert.equal(a.toString("hex"), b.toString("hex"));
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("v2 signature verification fails with wrong public key", async () => {
  const { dir } = await createTempExtensionDir();
  const keyA = generateEd25519KeyPair();
  const keyB = generateEd25519KeyPair();

  try {
    const pkg = await createExtensionPackageV2(dir, { privateKeyPem: keyA.privateKeyPem });
    assert.throws(() => verifyExtensionPackageV2(pkg, keyB.publicKeyPem), /signature verification failed/i);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("v2 checksum mismatch is detected", async () => {
  const { dir } = await createTempExtensionDir();
  const key = generateEd25519KeyPair();

  try {
    const pkg = await createExtensionPackageV2(dir, { privateKeyPem: key.privateKeyPem });
    const tampered = Buffer.from(pkg);
    const idx = tampered.indexOf(UNIQUE_MARKER);
    assert.ok(idx > 0, "marker should exist in package bytes");
    tampered[idx] ^= 0x01;

    assert.throws(() => verifyExtensionPackageV2(tampered, key.publicKeyPem), /checksum mismatch/i);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("v2 rejects path traversal entries", async () => {
  const { dir } = await createTempExtensionDir();
  const key = generateEd25519KeyPair();

  try {
    const pkg = await createExtensionPackageV2(dir, { privateKeyPem: key.privateKeyPem });
    const tampered = Buffer.from(pkg);

    const headerOffset = findTarHeaderOffset(tampered, ({ name }) => name === "files/package.json");
    assert.ok(typeof headerOffset === "number");

    const header = tampered.subarray(headerOffset, headerOffset + 512);
    header.fill(0, 0, 100);
    header.write("files/../evil.txt", 0, "utf8");
    updateTarChecksum(header);

    assert.throws(() => readExtensionPackageV2(tampered), /invalid path/i);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("marketplace store accepts v1 packages during transition", async (t) => {
  try {
    requireFromHere.resolve("sql.js");
  } catch {
    t.skip("sql.js dependency not installed in this environment");
    return;
  }

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-marketplace-v1-"));
  const dataDir = path.join(tmpRoot, "data");
  const { dir, manifest } = await createTempExtensionDir();

  try {
    const keys = generateEd25519KeyPair();

    const store = new MarketplaceStore({ dataDir });
    await store.init();
    await store.registerPublisher({
      publisher: manifest.publisher,
      tokenSha256: "ignored-for-unit-test",
      publicKeyPem: keys.publicKeyPem,
      verified: true,
    });

    const pkgBytes = await createExtensionPackageV1(dir);
    const signatureBase64 = signBytes(pkgBytes, keys.privateKeyPem);

    const published = await store.publishExtension({
      publisher: manifest.publisher,
      packageBytes: pkgBytes,
      signatureBase64,
    });

    assert.equal(published.id, `${manifest.publisher}.${manifest.name}`);
    assert.equal(published.version, manifest.version);

    const db = await store.db.getDb();
    const stmt = db.prepare(
      `SELECT format_version, file_count, unpacked_size, files_json
       FROM extension_versions WHERE extension_id = ? AND version = ? LIMIT 1`
    );
    stmt.bind([published.id, published.version]);
    assert.ok(stmt.step());
    const row = stmt.getAsObject();
    stmt.free();

    assert.equal(Number(row.format_version), 1);
    assert.ok(Number(row.file_count) > 0);
    assert.ok(Number(row.unpacked_size) > 0);
    assert.ok(Array.isArray(JSON.parse(String(row.files_json || "[]"))));
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    await fs.rm(dir, { recursive: true, force: true });
  }
});
