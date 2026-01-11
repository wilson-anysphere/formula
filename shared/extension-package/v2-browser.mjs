import { ed25519 } from "@noble/curves/ed25519";

import v2Core from "./core/v2-core.js";

const {
  PACKAGE_FORMAT_VERSION,
  canonicalJsonString,
  createSignaturePayloadBytes,
  decodeUtf8,
  isPlainObject,
  normalizePath,
  readExtensionPackageV2
} = v2Core;

const SIGNATURE_ALGORITHM = "ed25519";

function bytesToHex(bytes) {
  let out = "";
  for (const b of bytes) out += b.toString(16).padStart(2, "0");
  return out;
}

async function sha256Hex(bytes) {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle || typeof subtle.digest !== "function") {
    throw new Error("WebCrypto subtle.digest is required to verify extension packages in the browser");
  }

  const digest = await subtle.digest("SHA-256", bytes);
  return bytesToHex(new Uint8Array(digest));
}

function base64ToBytes(base64) {
  const raw = String(base64);
  if (typeof globalThis.atob === "function") {
    const bin = globalThis.atob(raw);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  // Node fallback (used by unit tests).
  if (typeof Buffer !== "undefined") {
    return Uint8Array.from(Buffer.from(raw, "base64"));
  }
  throw new Error("Base64 decoding is not available in this runtime");
}

function pemToDerBytes(pem) {
  const raw = String(pem)
    .trim()
    .replace(/-----BEGIN PUBLIC KEY-----/g, "")
    .replace(/-----END PUBLIC KEY-----/g, "")
    .replace(/\s+/g, "");
  if (!raw) throw new Error("Invalid public key PEM (empty)");
  return base64ToBytes(raw);
}

function readDerLength(bytes, offset) {
  const first = bytes[offset++];
  if ((first & 0x80) === 0) {
    return { length: first, offset };
  }

  const count = first & 0x7f;
  if (count === 0 || count > 4) {
    throw new Error("Invalid DER length encoding");
  }

  let length = 0;
  for (let i = 0; i < count; i++) {
    length = (length << 8) | bytes[offset++];
  }

  return { length, offset };
}

function expectDerTag(bytes, offset, tag, ctx) {
  const actual = bytes[offset];
  if (actual !== tag) {
    throw new Error(`Invalid DER: expected tag 0x${tag.toString(16)} for ${ctx} but got 0x${actual.toString(16)}`);
  }
  return offset + 1;
}

function parseEd25519PublicKeyFromSpki(spkiDer) {
  const bytes = spkiDer instanceof Uint8Array ? spkiDer : new Uint8Array(spkiDer);
  let offset = 0;

  offset = expectDerTag(bytes, offset, 0x30, "SubjectPublicKeyInfo");
  const seq = readDerLength(bytes, offset);
  offset = seq.offset;
  const seqEnd = offset + seq.length;
  if (seqEnd > bytes.length) throw new Error("Invalid DER: truncated SubjectPublicKeyInfo sequence");

  offset = expectDerTag(bytes, offset, 0x30, "AlgorithmIdentifier");
  const alg = readDerLength(bytes, offset);
  offset = alg.offset;
  const algEnd = offset + alg.length;
  if (algEnd > seqEnd) throw new Error("Invalid DER: truncated AlgorithmIdentifier sequence");

  offset = expectDerTag(bytes, offset, 0x06, "AlgorithmIdentifier.oid");
  const oidLen = readDerLength(bytes, offset);
  offset = oidLen.offset;
  const oidBytes = bytes.subarray(offset, offset + oidLen.length);
  offset += oidLen.length;

  // OID 1.3.101.112 (Ed25519) encodes to 06 03 2B 65 70.
  const expectedOid = [0x2b, 0x65, 0x70];
  if (oidBytes.length !== expectedOid.length) {
    throw new Error("Invalid Ed25519 public key: unexpected SPKI algorithm OID");
  }
  for (let i = 0; i < expectedOid.length; i++) {
    if (oidBytes[i] !== expectedOid[i]) {
      throw new Error("Invalid Ed25519 public key: unexpected SPKI algorithm OID");
    }
  }

  // Ed25519 AlgorithmIdentifier has no params; skip any remaining bytes in the sequence.
  offset = algEnd;

  offset = expectDerTag(bytes, offset, 0x03, "SubjectPublicKeyInfo.publicKey");
  const bitStringLen = readDerLength(bytes, offset);
  offset = bitStringLen.offset;
  if (offset + bitStringLen.length > seqEnd) {
    throw new Error("Invalid DER: truncated public key bit string");
  }

  const unusedBits = bytes[offset++];
  if (unusedBits !== 0) {
    throw new Error("Invalid Ed25519 public key: expected 0 unused bits in BIT STRING");
  }

  const keyBytes = bytes.subarray(offset, offset + (bitStringLen.length - 1));
  if (keyBytes.length !== 32) {
    throw new Error(`Invalid Ed25519 public key: expected 32 bytes but got ${keyBytes.length}`);
  }
  return keyBytes;
}

async function verifyEd25519Signature(payloadBytes, signatureBase64, publicKeyPem) {
  const signature = base64ToBytes(signatureBase64);

  const subtle = globalThis.crypto?.subtle;
  if (subtle && typeof subtle.importKey === "function" && typeof subtle.verify === "function") {
    try {
      const key = await subtle.importKey(
        "spki",
        pemToDerBytes(publicKeyPem),
        { name: "Ed25519" },
        false,
        ["verify"]
      );
      const ok = await subtle.verify({ name: "Ed25519" }, key, signature, payloadBytes);
      return Boolean(ok);
    } catch {
      // fall through to JS verifier
    }
  }

  const publicKeyBytes = parseEd25519PublicKeyFromSpki(pemToDerBytes(publicKeyPem));
  try {
    return ed25519.verify(signature, payloadBytes, publicKeyBytes);
  } catch {
    return false;
  }
}

export async function verifyExtensionPackageV2Browser(packageBytes, publicKeyPem) {
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
  const ok = await verifyEd25519Signature(payloadBytes, signature.signatureBase64, publicKeyPem);
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
    packageJson = JSON.parse(decodeUtf8(packageJsonBytes));
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

    // eslint-disable-next-line no-await-in-loop
    const actualSha = await sha256Hex(data);
    if (actualSha !== expected.sha256) {
      throw new Error(`Checksum mismatch for ${relPath}`);
    }
    if (data.length !== expected.size) {
      throw new Error(`Size mismatch for ${relPath}`);
    }

    unpackedSize += data.length;
    if (!readme && relPath.toLowerCase() === "readme.md") {
      readme = decodeUtf8(data);
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
    readme
  };
}

export { PACKAGE_FORMAT_VERSION, readExtensionPackageV2 };

