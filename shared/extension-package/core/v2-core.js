const PACKAGE_FORMAT_VERSION = 2;

const TAR_BLOCK_SIZE = 512;
const MAX_TAR_ENTRIES = 10_000;
const MAX_PACKAGE_FILES = 5_000;
const MAX_MANIFEST_JSON_BYTES = 1 * 1024 * 1024;
const MAX_CHECKSUMS_JSON_BYTES = 4 * 1024 * 1024;
const MAX_SIGNATURE_JSON_BYTES = 64 * 1024;

// Cache codec instances: TextEncoder/TextDecoder are small but allocating per call is wasteful.
const textEncoder = typeof TextEncoder === "undefined" ? null : new TextEncoder();
const utf8Decoder = typeof TextDecoder === "undefined" ? null : new TextDecoder("utf-8", { fatal: false });

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

function createSignaturePayloadBytes(manifest, checksums) {
  if (typeof TextEncoder === "undefined") {
    throw new Error("TextEncoder is required to verify v2 extension packages");
  }
  return textEncoder.encode(canonicalJsonString({ manifest, checksums }));
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

function decodeNullTerminatedUtf8(bytes) {
  if (typeof TextDecoder === "undefined") {
    throw new Error("TextDecoder is required to verify v2 extension packages");
  }
  const idx = bytes.indexOf(0);
  const slice = idx === -1 ? bytes : bytes.subarray(0, idx);
  return utf8Decoder.decode(slice);
}

function decodeUtf8(bytes) {
  if (typeof TextDecoder === "undefined") {
    throw new Error("TextDecoder is required to verify v2 extension packages");
  }
  return utf8Decoder.decode(bytes);
}

function parseTarString(block, offset, length) {
  const raw = block.subarray(offset, offset + length);
  return decodeNullTerminatedUtf8(raw);
}

function parseTarOctal(block, offset, length) {
  const raw = block.subarray(offset, offset + length);
  let str = "";
  for (let i = 0; i < raw.length; i++) {
    const b = raw[i];
    if (b === 0) break;
    str += String.fromCharCode(b);
  }
  str = str.trim();
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
      const remainingBytes = archiveBytes.length - offset;
      if (remainingBytes < TAR_BLOCK_SIZE) {
        throw new Error("Invalid tar archive: missing end-of-archive marker");
      }

      // Many tar implementations pad archives to 10KB "record" boundaries. Allow some trailing
      // zero padding, but cap it to prevent attackers from appending huge unused payloads.
      const maxTrailerBytes = TAR_BLOCK_SIZE * 20; // 10240 bytes
      if (remainingBytes > maxTrailerBytes) {
        throw new Error("Invalid tar archive: excessive trailing padding");
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

function readExtensionPackageV2(packageBytes) {
  const required = {
    manifest: null,
    checksums: null,
    signature: null
  };

  const payloadFiles = new Map();
  const payloadFilesFolded = new Set();

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
      required.manifest = JSON.parse(decodeUtf8(entry.data));
      continue;
    }
    if (name === "checksums.json") {
      if (required.checksums) {
        throw new Error("Invalid extension package: duplicate checksums.json");
      }
      if (entry.data.length > MAX_CHECKSUMS_JSON_BYTES) {
        throw new Error("Invalid extension package: checksums.json is too large");
      }
      required.checksums = JSON.parse(decodeUtf8(entry.data));
      continue;
    }
    if (name === "signature.json") {
      if (required.signature) {
        throw new Error("Invalid extension package: duplicate signature.json");
      }
      if (entry.data.length > MAX_SIGNATURE_JSON_BYTES) {
        throw new Error("Invalid extension package: signature.json is too large");
      }
      required.signature = JSON.parse(decodeUtf8(entry.data));
      continue;
    }

    if (name.startsWith("files/")) {
      if (payloadFiles.size >= MAX_PACKAGE_FILES) {
        throw new Error(`Invalid extension package: too many files (>${MAX_PACKAGE_FILES})`);
      }
      const relPath = normalizePath(name.slice("files/".length));
      const folded = relPath.toLowerCase();
      if (payloadFilesFolded.has(folded)) {
        throw new Error(`Invalid extension package: duplicate file path (case-insensitive): ${relPath}`);
      }
      if (payloadFiles.has(relPath)) throw new Error(`Duplicate file in package: ${relPath}`);
      payloadFiles.set(relPath, entry.data);
      payloadFilesFolded.add(folded);
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
    files: payloadFiles
  };
}

module.exports = {
  PACKAGE_FORMAT_VERSION,
  TAR_BLOCK_SIZE,
  MAX_TAR_ENTRIES,
  MAX_PACKAGE_FILES,
  MAX_MANIFEST_JSON_BYTES,
  MAX_CHECKSUMS_JSON_BYTES,
  MAX_SIGNATURE_JSON_BYTES,
  canonicalJsonString,
  createSignaturePayloadBytes,
  decodeUtf8,
  isPlainObject,
  iterateTarEntries,
  normalizePath,
  readExtensionPackageV2
};
