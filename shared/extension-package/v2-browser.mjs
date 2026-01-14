import * as v2CoreImport from "./core/v2-core.mjs";

// `v2-core.js` is authored as CommonJS (shared with Node tooling). When served through Vite it may
// expose only named exports (no `default` export). Normalize to a single object so browser builds
// can consume it reliably.
const v2Core = v2CoreImport.default ?? v2CoreImport;

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
// NOTE: Keep these limits in sync with the desktop verifier in
// `apps/desktop/src-tauri/src/ed25519_verifier.rs`.
const MAX_SIGNATURE_PAYLOAD_BYTES = 5 * 1024 * 1024; // 5MB
const MAX_PUBLIC_KEY_PEM_BYTES = 64 * 1024; // 64KB
const MAX_SIGNATURE_BASE64_BYTES = 1024;

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

function isEd25519NotSupportedError(error) {
  const name = typeof error?.name === "string" ? error.name : "";
  if (name === "NotSupportedError") return true;

  // Best-effort compatibility for runtimes that throw a generic error with a message instead.
  const message = String(error?.message ?? "");
  if (!message) return false;

  return /ed25519/i.test(message) && /not supported|unsupported/i.test(message);
}

function getTauriInvoke() {
  try {
    const tauri = globalThis.__TAURI__;
    const invoke = tauri?.core?.invoke;
    return typeof invoke === "function" ? invoke : null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so browser verification can fall back cleanly.
    return null;
  }
}

async function verifyEd25519Signature(payloadBytes, signatureBase64, publicKeyPem) {
  if (payloadBytes.length > MAX_SIGNATURE_PAYLOAD_BYTES) {
    throw new Error(`Signature payload is too large (max ${MAX_SIGNATURE_PAYLOAD_BYTES} bytes)`);
  }
  if (String(publicKeyPem).length > MAX_PUBLIC_KEY_PEM_BYTES) {
    throw new Error(`Public key PEM is too large (max ${MAX_PUBLIC_KEY_PEM_BYTES} bytes)`);
  }
  if (String(signatureBase64).length > MAX_SIGNATURE_BASE64_BYTES) {
    throw new Error(`Signature base64 is too large (max ${MAX_SIGNATURE_BASE64_BYTES} bytes)`);
  }

  const signature = base64ToBytes(signatureBase64);

  const subtle = globalThis.crypto?.subtle;
  try {
    if (!subtle || typeof subtle.importKey !== "function" || typeof subtle.verify !== "function") {
      throw new Error(
        "WebCrypto (crypto.subtle.importKey()/verify()) is required to verify extension packages in the browser"
      );
    }

    const key = await subtle.importKey(
      "spki",
      pemToDerBytes(publicKeyPem),
      { name: "Ed25519" },
      false,
      ["verify"]
    );

    const ok = await subtle.verify({ name: "Ed25519" }, key, signature, payloadBytes);
    return Boolean(ok);
  } catch (error) {
    const message = String(error?.message ?? error);
    if (isEd25519NotSupportedError(error)) {
      const invoke = getTauriInvoke();
      if (invoke) {
        try {
          const ok = await invoke("verify_ed25519_signature", {
            payload: Array.from(payloadBytes),
            signature_base64: String(signatureBase64),
            public_key_pem: String(publicKeyPem)
          });
          return Boolean(ok);
        } catch (invokeError) {
          const invokeMsg = String(invokeError?.message ?? invokeError);
          throw new Error(`Failed to verify extension signature via Tauri IPC: ${invokeMsg}`);
        }
      }
      throw new Error(
        "This environment's WebCrypto implementation does not support Ed25519, so extension packages cannot be verified. " +
          "If you are running Formula Desktop, ensure the native verifier is available. " +
          `Original error: ${message}`
      );
    }
    throw new Error(`Failed to verify extension signature via WebCrypto: ${message}`);
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
