const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const path = require("node:path");

const { SIGNATURE_ALGORITHM, signBytes, verifyBytesSignature } = require("../crypto/signing");

const v2Core = require("./core/v2-core");
const {
  PACKAGE_FORMAT_VERSION,
  TAR_BLOCK_SIZE,
  canonicalJsonString,
  createSignaturePayloadBytes: createSignaturePayloadBytesCore,
  isPlainObject,
  normalizePath,
  readExtensionPackageV2: readExtensionPackageV2Core,
} = v2Core;

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function canonicalJsonBytes(value) {
  return Buffer.from(canonicalJsonString(value), "utf8");
}

async function walkFiles(rootDir) {
  const results = [];

  async function visit(dir) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.isSymbolicLink()) continue;

      const abs = path.join(dir, entry.name);
      const rel = path.relative(rootDir, abs).replace(/\\/g, "/");

      if (rel === "" || rel.startsWith("..") || path.isAbsolute(rel)) continue;

      if (entry.isDirectory()) {
        if (
          entry.name === "node_modules" ||
          entry.name === ".git" ||
          entry.name === "__tests__" ||
          entry.name === "test" ||
          entry.name === "tests"
        ) {
          continue;
        }
        await visit(abs);
        continue;
      }

      if (!entry.isFile()) continue;

      const lowerName = entry.name.toLowerCase();
      if (
        lowerName.endsWith(".test.js") ||
        lowerName.endsWith(".test.cjs") ||
        lowerName.endsWith(".test.mjs") ||
        lowerName.endsWith(".spec.js") ||
        lowerName.endsWith(".spec.cjs") ||
        lowerName.endsWith(".spec.mjs")
      ) {
        continue;
      }
      results.push({ abs, rel: normalizePath(rel) });
    }
  }

  await visit(rootDir);
  results.sort((a, b) => (a.rel < b.rel ? -1 : a.rel > b.rel ? 1 : 0));
  return results;
}

async function loadExtensionManifest(extensionDir) {
  const manifestPath = path.join(extensionDir, "package.json");
  const raw = await fs.readFile(manifestPath, "utf8");
  return JSON.parse(raw);
}

function createSignaturePayloadBytes(manifest, checksums) {
  return Buffer.from(createSignaturePayloadBytesCore(manifest, checksums));
}

function encodeTarOctal(buffer, offset, length, value) {
  const str = value.toString(8);
  const maxDigits = length - 1;
  if (str.length > maxDigits) {
    throw new Error(`Tar field overflow: ${value} does not fit in ${length} bytes`);
  }
  buffer.fill(0, offset, offset + length);
  buffer.write(str.padStart(maxDigits, "0"), offset, maxDigits, "ascii");
  buffer[offset + length - 1] = 0;
}

function encodeTarString(buffer, offset, length, value) {
  buffer.fill(0, offset, offset + length);
  if (!value) return;
  const bytes = Buffer.from(String(value), "utf8");
  if (bytes.length > length) {
    throw new Error(`Tar field overflow: string is ${bytes.length} bytes (max ${length})`);
  }
  bytes.copy(buffer, offset);
}

function splitTarPath(name) {
  const bytes = Buffer.from(name, "utf8");
  if (bytes.length <= 100) return { name, prefix: "" };

  const parts = name.split("/");
  for (let i = parts.length - 1; i > 0; i--) {
    const prefix = parts.slice(0, i).join("/");
    const suffix = parts.slice(i).join("/");
    if (Buffer.byteLength(suffix, "utf8") <= 100 && Buffer.byteLength(prefix, "utf8") <= 155) {
      return { name: suffix, prefix };
    }
  }

  throw new Error(`Tar path too long: ${name}`);
}

function createTarHeader({ name, size, typeflag = "0", mode = 0o644 }) {
  const header = Buffer.alloc(TAR_BLOCK_SIZE, 0);
  const split = splitTarPath(name);

  encodeTarString(header, 0, 100, split.name);
  encodeTarOctal(header, 100, 8, mode);
  encodeTarOctal(header, 108, 8, 0); // uid
  encodeTarOctal(header, 116, 8, 0); // gid
  encodeTarOctal(header, 124, 12, size);
  encodeTarOctal(header, 136, 12, 0); // mtime for determinism

  // checksum placeholder (spaces)
  header.fill(0x20, 148, 156);

  encodeTarString(header, 156, 1, typeflag);
  encodeTarString(header, 257, 6, "ustar\0");
  encodeTarString(header, 263, 2, "00");
  encodeTarString(header, 345, 155, split.prefix);

  let sum = 0;
  for (const b of header) sum += b;
  const checksumStr = sum.toString(8).padStart(6, "0");
  encodeTarString(header, 148, 6, checksumStr);
  header[154] = 0;
  header[155] = 0x20;

  return header;
}

function createTarArchive(entries) {
  const chunks = [];
  for (const entry of entries) {
    const data = entry.data ?? Buffer.alloc(0);
    const header = createTarHeader({ name: entry.name, size: data.length });
    chunks.push(header, data);

    const pad = (TAR_BLOCK_SIZE - (data.length % TAR_BLOCK_SIZE)) % TAR_BLOCK_SIZE;
    if (pad) chunks.push(Buffer.alloc(pad, 0));
  }

  // End-of-archive: two empty blocks.
  chunks.push(Buffer.alloc(TAR_BLOCK_SIZE * 2, 0));
  return Buffer.concat(chunks);
}

function safeJoin(baseDir, relPath) {
  const normalized = normalizePath(relPath);
  const full = path.join(baseDir, normalized);
  const relative = path.relative(baseDir, full);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Path traversal in extension package: ${relPath}`);
  }
  return full;
}

async function createExtensionPackageV2(extensionDir, { privateKeyPem } = {}) {
  if (!privateKeyPem) {
    throw new Error("privateKeyPem is required to create a signed v2 extension package");
  }

  const manifest = await loadExtensionManifest(extensionDir);
  const manifestBytes = canonicalJsonBytes(manifest);

  const entries = await walkFiles(extensionDir);
  const files = [];
  for (const entry of entries) {
    let data = await fs.readFile(entry.abs);
    if (entry.rel === "package.json") {
      data = manifestBytes;
    }
    files.push({
      path: entry.rel,
      data,
      size: data.length,
      sha256: sha256Hex(data),
    });
  }

  const checksums = {
    algorithm: "sha256",
    files: Object.fromEntries(
      files.map((f) => [f.path, { sha256: f.sha256, size: f.size }]),
    ),
  };
  const checksumsBytes = canonicalJsonBytes(checksums);

  const payloadBytes = createSignaturePayloadBytes(manifest, checksums);
  const signatureBase64 = signBytes(payloadBytes, privateKeyPem, { algorithm: SIGNATURE_ALGORITHM });
  const signature = {
    algorithm: SIGNATURE_ALGORITHM,
    formatVersion: PACKAGE_FORMAT_VERSION,
    signatureBase64,
  };
  const signatureBytes = canonicalJsonBytes(signature);

  // Deterministic entry ordering for streaming:
  // - signature material first (manifest/checksums/signature)
  // - then payload files, sorted lexicographically by path.
  const tarEntries = [
    { name: "manifest.json", data: manifestBytes },
    { name: "checksums.json", data: checksumsBytes },
    { name: "signature.json", data: signatureBytes },
    ...files.map((f) => ({ name: `files/${f.path}`, data: f.data })),
  ];

  return createTarArchive(tarEntries);
}

function readExtensionPackageV2(packageBytes) {
  return readExtensionPackageV2Core(packageBytes);
}

function verifyExtensionPackageV2Parsed(parsed, publicKeyPem) {
  const { manifest, checksums, signature, files } = parsed;
  if (!isPlainObject(manifest)) throw new Error("Invalid manifest.json (expected object)");
  if (!isPlainObject(checksums) || checksums.algorithm !== "sha256" || !isPlainObject(checksums.files)) {
    throw new Error("Invalid checksums.json");
  }
  if (
    !isPlainObject(signature) ||
    signature.algorithm !== SIGNATURE_ALGORITHM ||
    signature.formatVersion !== PACKAGE_FORMAT_VERSION ||
    typeof signature.signatureBase64 !== "string"
  ) {
    throw new Error("Invalid signature.json");
  }

  const MAX_CHECKSUM_ENTRIES = 5_000;
  const checksumEntries = new Map();
  let checksumCount = 0;
  for (const [rawPath, entry] of Object.entries(checksums.files)) {
    checksumCount += 1;
    if (checksumCount > MAX_CHECKSUM_ENTRIES) {
      throw new Error(`checksums.json contains too many entries (>${MAX_CHECKSUM_ENTRIES})`);
    }
    const normalized = normalizePath(rawPath);
    if (normalized !== rawPath) {
      throw new Error(`checksums.json contains non-normalized path: ${rawPath}`);
    }
    if (!isPlainObject(entry)) {
      throw new Error(`checksums.json entry for ${rawPath} must be an object`);
    }
    const sha = typeof entry.sha256 === "string" ? entry.sha256.toLowerCase() : null;
    if (!sha || !/^[0-9a-f]{64}$/.test(sha)) {
      throw new Error(`checksums.json entry for ${rawPath} has invalid sha256`);
    }
    const size = entry.size;
    if (
      typeof size !== "number" ||
      !Number.isFinite(size) ||
      size < 0 ||
      !Number.isInteger(size) ||
      size > Number.MAX_SAFE_INTEGER
    ) {
      throw new Error(`checksums.json entry for ${rawPath} has invalid size`);
    }
    checksumEntries.set(rawPath, { sha256: sha, size });
  }

  const payloadBytes = createSignaturePayloadBytes(manifest, checksums);
  const ok = verifyBytesSignature(payloadBytes, signature.signatureBase64, publicKeyPem, { algorithm: SIGNATURE_ALGORITHM });
  if (!ok) throw new Error("Package signature verification failed");

  /** @type {{path: string, sha256: string, size: number}[]} */
  const fileRecords = [];
  let unpackedSize = 0;
  let readme = "";

  const packageJsonBytes = files.get("package.json");
  if (!packageJsonBytes) {
    throw new Error("Invalid extension package: missing files/package.json");
  }
  let packageJson = null;
  try {
    packageJson = JSON.parse(packageJsonBytes.toString("utf8"));
  } catch {
    throw new Error("Invalid files/package.json (expected JSON)");
  }
  if (!isPlainObject(packageJson)) {
    throw new Error("Invalid files/package.json (expected object)");
  }
  if (canonicalJsonString(packageJson) !== canonicalJsonString(manifest)) {
    throw new Error("files/package.json does not match manifest.json");
  }

  const expectedPaths = new Set(checksumEntries.keys());

  for (const [relPath, data] of files.entries()) {
    const expected = checksumEntries.get(relPath);
    if (!expected) {
      throw new Error(`checksums.json missing entry for ${relPath}`);
    }

    const actualSha = sha256Hex(data);
    if (actualSha !== expected.sha256) {
      throw new Error(`Checksum mismatch for ${relPath}`);
    }
    if (data.length !== expected.size) {
      throw new Error(`Size mismatch for ${relPath}`);
    }

    unpackedSize += data.length;
    if (!readme && relPath.toLowerCase() === "readme.md") {
      readme = data.toString("utf8");
    }
    expectedPaths.delete(relPath);
    fileRecords.push({ path: relPath, sha256: actualSha, size: data.length });
  }

  if (expectedPaths.size > 0) {
    throw new Error(`checksums.json contains entries missing from archive: ${[...expectedPaths].join(", ")}`);
  }

  fileRecords.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));

  return {
    manifest,
    signatureBase64: signature.signatureBase64,
    files: fileRecords,
    unpackedSize,
    fileCount: fileRecords.length,
    readme,
  };
}

function verifyExtensionPackageV2(packageBytes, publicKeyPem) {
  const parsed = readExtensionPackageV2(packageBytes);
  return verifyExtensionPackageV2Parsed(parsed, publicKeyPem);
}

async function extractExtensionPackageV2FromParsed(parsed, destDir) {
  await fs.mkdir(destDir, { recursive: true });

  for (const [relPath, data] of parsed.files.entries()) {
    const outPath = safeJoin(destDir, relPath);
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    await fs.writeFile(outPath, data);
  }
}

async function extractExtensionPackageV2(packageBytes, destDir) {
  const parsed = readExtensionPackageV2(packageBytes);
  await extractExtensionPackageV2FromParsed(parsed, destDir);
  return parsed.manifest;
}

module.exports = {
  PACKAGE_FORMAT_VERSION,
  canonicalJsonBytes,
  createSignaturePayloadBytes,
  createExtensionPackageV2,
  extractExtensionPackageV2FromParsed,
  extractExtensionPackageV2,
  loadExtensionManifest,
  readExtensionPackageV2,
  verifyExtensionPackageV2,
  verifyExtensionPackageV2Parsed,
};
