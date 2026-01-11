const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const path = require("node:path");

const { SIGNATURE_ALGORITHM, signBytes, verifyBytesSignature } = require("../crypto/signing");

const PACKAGE_FORMAT_VERSION = 2;

const TAR_BLOCK_SIZE = 512;
const MAX_TAR_ENTRIES = 10_000;
const MAX_PACKAGE_FILES = 5_000;
const MAX_MANIFEST_JSON_BYTES = 1 * 1024 * 1024;
const MAX_CHECKSUMS_JSON_BYTES = 4 * 1024 * 1024;
const MAX_SIGNATURE_JSON_BYTES = 64 * 1024;

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

// Deterministic JSON stringify with lexicographically sorted object keys.
function canonicalJsonString(value) {
  if (value === null) return "null";
  if (typeof value === "string") return JSON.stringify(value);
  if (typeof value === "number") return JSON.stringify(value);
  if (typeof value === "boolean") return value ? "true" : "false";
  if (Array.isArray(value)) {
    return `[${value.map((v) => canonicalJsonString(v)).join(",")}]`;
  }
  if (isPlainObject(value)) {
    const keys = Object.keys(value).sort();
    const items = keys
      .filter((k) => value[k] !== undefined)
      .map((k) => `${JSON.stringify(k)}:${canonicalJsonString(value[k])}`);
    return `{${items.join(",")}}`;
  }
  // JSON doesn't support functions, symbols, BigInt, etc.
  return JSON.stringify(value);
}

function canonicalJsonBytes(value) {
  return Buffer.from(canonicalJsonString(value), "utf8");
}

function normalizePath(relPath) {
  const normalized = String(relPath).replace(/\\/g, "/");
  if (normalized.startsWith("/") || normalized.includes("\0")) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  const parts = normalized.split("/");
  if (parts.some((p) => p === "" || p === "." || p === "..")) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  // Disallow ':' to avoid Windows drive-relative / alternate stream paths (and keep packages portable).
  if (parts.some((p) => p.includes(":"))) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  // Cross-platform safety: Windows strips trailing dots/spaces, and also reserves certain device names.
  const windowsReservedRe = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;
  for (const part of parts) {
    // Windows disallows these characters in file/directory names.
    if (/[<>:"|?*]/.test(part)) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
    if (part.endsWith(" ") || part.endsWith(".")) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
    if (windowsReservedRe.test(part)) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
  }
  return parts.join("/");
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
        if (entry.name === "node_modules" || entry.name === ".git") continue;
        await visit(abs);
        continue;
      }

      if (!entry.isFile()) continue;
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
  return canonicalJsonBytes({ manifest, checksums });
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

function parseTarString(block, offset, length) {
  const raw = block.subarray(offset, offset + length);
  const idx = raw.indexOf(0);
  const slice = idx === -1 ? raw : raw.subarray(0, idx);
  return slice.toString("utf8");
}

function parseTarOctal(block, offset, length) {
  const raw = block.subarray(offset, offset + length);
  const str = raw.toString("ascii").replace(/\0.*$/, "").trim();
  if (!str) return 0;
  return Number.parseInt(str, 8);
}

function verifyTarChecksum(block) {
  const expected = parseTarOctal(block, 148, 8);
  let sum = 0;
  for (let i = 0; i < TAR_BLOCK_SIZE; i++) {
    const val = i >= 148 && i < 156 ? 0x20 : block[i];
    sum += val;
  }
  return sum === expected;
}

function* iterateTarEntries(archiveBytes) {
  if (archiveBytes.length % TAR_BLOCK_SIZE !== 0) {
    throw new Error("Invalid tar archive length (expected 512-byte blocks)");
  }
  let offset = 0;
  let sawEndMarker = false;
  while (offset + TAR_BLOCK_SIZE <= archiveBytes.length) {
    const header = archiveBytes.subarray(offset, offset + TAR_BLOCK_SIZE);
    offset += TAR_BLOCK_SIZE;

    // EOF blocks.
    if (header.every((b) => b === 0)) {
      sawEndMarker = true;

      // v2 packages should terminate with at least two 512-byte zero blocks. We already consumed
      // the first one as `header`, so require at least one more full block remaining.
      if (archiveBytes.length - offset < TAR_BLOCK_SIZE) {
        throw new Error("Invalid tar archive: missing end-of-archive marker");
      }

      // Ensure there is no trailing non-zero data after the end-of-archive marker; this prevents
      // attackers from smuggling large unused payloads past size validation.
      for (let i = offset; i < archiveBytes.length; i++) {
        if (archiveBytes[i] !== 0) {
          throw new Error("Invalid tar archive: unexpected data after end-of-archive marker");
        }
      }
      break;
    }

    if (!verifyTarChecksum(header)) {
      throw new Error("Invalid tar checksum");
    }

    const name = parseTarString(header, 0, 100);
    const prefix = parseTarString(header, 345, 155);
    const fullName = prefix ? `${prefix}/${name}` : name;
    const size = parseTarOctal(header, 124, 12);
    const typeflag = parseTarString(header, 156, 1) || "0";

    if (!Number.isFinite(size) || size < 0) {
      throw new Error(`Invalid tar entry size: ${fullName}`);
    }

    const dataStart = offset;
    const dataEnd = dataStart + size;
    if (dataEnd > archiveBytes.length) {
      throw new Error("Truncated tar entry data");
    }
    const data = archiveBytes.subarray(dataStart, dataEnd);

    offset = dataStart + Math.ceil(size / TAR_BLOCK_SIZE) * TAR_BLOCK_SIZE;

    yield { name: fullName, size, typeflag, data };
  }

  if (!sawEndMarker) {
    throw new Error("Invalid tar archive: missing end-of-archive marker");
  }
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
  const required = {
    manifest: null,
    checksums: null,
    signature: null,
  };

  const payloadFiles = new Map();

  let entriesSeen = 0;
  for (const entry of iterateTarEntries(packageBytes)) {
    entriesSeen += 1;
    if (entriesSeen > MAX_TAR_ENTRIES) {
      throw new Error(`Invalid extension package: too many tar entries (>${MAX_TAR_ENTRIES})`);
    }

    if (entry.typeflag !== "0" && entry.typeflag !== "\0") {
      if (entry.typeflag === "5") continue; // directory entry
      throw new Error(`Unsupported tar entry type: ${entry.typeflag} (${entry.name})`);
    }

    const normalizedName = entry.name.endsWith("/") ? entry.name.slice(0, -1) : entry.name;
    const name = normalizePath(normalizedName);

    if (name === "manifest.json") {
      if (required.manifest) {
        throw new Error("Invalid extension package: duplicate manifest.json");
      }
      if (entry.data.length > MAX_MANIFEST_JSON_BYTES) {
        throw new Error("Invalid extension package: manifest.json is too large");
      }
      required.manifest = JSON.parse(entry.data.toString("utf8"));
      continue;
    }
    if (name === "checksums.json") {
      if (required.checksums) {
        throw new Error("Invalid extension package: duplicate checksums.json");
      }
      if (entry.data.length > MAX_CHECKSUMS_JSON_BYTES) {
        throw new Error("Invalid extension package: checksums.json is too large");
      }
      required.checksums = JSON.parse(entry.data.toString("utf8"));
      continue;
    }
    if (name === "signature.json") {
      if (required.signature) {
        throw new Error("Invalid extension package: duplicate signature.json");
      }
      if (entry.data.length > MAX_SIGNATURE_JSON_BYTES) {
        throw new Error("Invalid extension package: signature.json is too large");
      }
      required.signature = JSON.parse(entry.data.toString("utf8"));
      continue;
    }

    if (name.startsWith("files/")) {
      if (payloadFiles.size >= MAX_PACKAGE_FILES) {
        throw new Error(`Invalid extension package: too many files (>${MAX_PACKAGE_FILES})`);
      }
      const relPath = normalizePath(name.slice("files/".length));
      if (payloadFiles.has(relPath)) throw new Error(`Duplicate file in package: ${relPath}`);
      payloadFiles.set(relPath, entry.data);
      continue;
    }

    throw new Error(`Unexpected entry in extension package: ${name}`);
  }

  if (!required.manifest || !required.checksums || !required.signature) {
    throw new Error("Invalid extension package: missing required manifest/checksums/signature");
  }

  return {
    formatVersion: PACKAGE_FORMAT_VERSION,
    manifest: required.manifest,
    checksums: required.checksums,
    signature: required.signature,
    files: payloadFiles,
  };
}

function verifyExtensionPackageV2(packageBytes, publicKeyPem) {
  const parsed = readExtensionPackageV2(packageBytes);

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
  };
}

async function extractExtensionPackageV2(packageBytes, destDir) {
  const parsed = readExtensionPackageV2(packageBytes);
  await fs.mkdir(destDir, { recursive: true });

  for (const [relPath, data] of parsed.files.entries()) {
    const outPath = safeJoin(destDir, relPath);
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    await fs.writeFile(outPath, data);
  }

  return parsed.manifest;
}

module.exports = {
  PACKAGE_FORMAT_VERSION,
  canonicalJsonBytes,
  createSignaturePayloadBytes,
  createExtensionPackageV2,
  extractExtensionPackageV2,
  loadExtensionManifest,
  readExtensionPackageV2,
  verifyExtensionPackageV2,
};
