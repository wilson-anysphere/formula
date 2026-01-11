const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const { createRequire } = require("node:module");
const os = require("node:os");
const path = require("node:path");

const {
  createExtensionPackageV1,
  createExtensionPackageV2,
  canonicalJsonBytes,
  createSignaturePayloadBytes,
  readExtensionPackageV2,
  verifyExtensionPackageV2,
} = require("../../../../shared/extension-package");
const { generateEd25519KeyPair, sha256, signBytes } = require("../../../../shared/crypto/signing");
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

function writeTarOctal(header, offset, length, value) {
  const str = value.toString(8);
  const maxDigits = length - 1;
  header.fill(0, offset, offset + length);
  header.write(str.padStart(maxDigits, "0"), offset, maxDigits, "ascii");
  header[offset + length - 1] = 0;
}

function writeTarString(header, offset, length, value) {
  header.fill(0, offset, offset + length);
  if (!value) return;
  const bytes = Buffer.from(String(value), "utf8");
  bytes.copy(header, offset, 0, Math.min(bytes.length, length));
}

function createTarHeader(name, size) {
  const header = Buffer.alloc(512, 0);
  writeTarString(header, 0, 100, name);
  writeTarOctal(header, 100, 8, 0o644);
  writeTarOctal(header, 108, 8, 0);
  writeTarOctal(header, 116, 8, 0);
  writeTarOctal(header, 124, 12, size);
  writeTarOctal(header, 136, 12, 0);
  writeTarString(header, 156, 1, "0");
  writeTarString(header, 257, 6, "ustar");
  updateTarChecksum(header);
  return header;
}

function createTarArchive(entries) {
  const chunks = [];
  for (const entry of entries) {
    const data = entry.data ?? Buffer.alloc(0);
    chunks.push(createTarHeader(entry.name, data.length), data);
    const pad = (512 - (data.length % 512)) % 512;
    if (pad) chunks.push(Buffer.alloc(pad, 0));
  }
  chunks.push(Buffer.alloc(1024, 0));
  return Buffer.concat(chunks);
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

test("v2 rejects packages where files/package.json differs from manifest.json", () => {
  const key = generateEd25519KeyPair();

  const manifest = {
    name: "temp-ext",
    publisher: "temp-pub",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  };

  const packageJson = {
    ...manifest,
    name: "different-ext",
  };

  const jsBytes = Buffer.from("module.exports = { activate() {} };\n", "utf8");
  const packageJsonBytes = canonicalJsonBytes(packageJson);

  const checksums = {
    algorithm: "sha256",
    files: {
      "dist/extension.js": { sha256: sha256(jsBytes), size: jsBytes.length },
      "package.json": { sha256: sha256(packageJsonBytes), size: packageJsonBytes.length },
    },
  };

  const signaturePayload = createSignaturePayloadBytes(manifest, checksums);
  const signatureBase64 = signBytes(signaturePayload, key.privateKeyPem);
  const signatureBytes = canonicalJsonBytes({
    algorithm: "ed25519",
    formatVersion: 2,
    signatureBase64,
  });

  const archive = createTarArchive([
    { name: "manifest.json", data: canonicalJsonBytes(manifest) },
    { name: "checksums.json", data: canonicalJsonBytes(checksums) },
    { name: "signature.json", data: signatureBytes },
    { name: "files/dist/extension.js", data: jsBytes },
    { name: "files/package.json", data: packageJsonBytes },
  ]);

  assert.throws(() => verifyExtensionPackageV2(archive, key.publicKeyPem), /does not match manifest/i);
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
