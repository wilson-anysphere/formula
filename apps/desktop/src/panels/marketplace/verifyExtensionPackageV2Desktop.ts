import type { VerifiedExtensionPackageV2 } from "../../../../../shared/extension-package/v2-browser.mjs";
import v2Core from "../../../../../shared/extension-package/core/v2-core.mjs";

const {
  PACKAGE_FORMAT_VERSION,
  canonicalJsonString,
  createSignaturePayloadBytes,
  decodeUtf8,
  isPlainObject,
  normalizePath,
  readExtensionPackageV2
} = v2Core as any;

const SIGNATURE_ALGORITHM = "ed25519";

function bytesToHex(bytes: Uint8Array): string {
  let out = "";
  for (const b of bytes) out += b.toString(16).padStart(2, "0");
  return out;
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle || typeof subtle.digest !== "function") {
    throw new Error("WebCrypto subtle.digest is required to verify extension packages");
  }

  // `crypto.subtle.digest` expects a BufferSource backed by an `ArrayBuffer`. TypeScript models
  // `Uint8Array` as potentially backed by a `SharedArrayBuffer` (`ArrayBufferLike`), so normalize
  // to an `ArrayBuffer`-backed view for type safety.
  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  const digest = await subtle.digest("SHA-256", normalized);
  return bytesToHex(new Uint8Array(digest));
}

function base64ToBytes(base64: string): Uint8Array {
  const raw = String(base64);
  if (typeof globalThis.atob === "function") {
    const bin = globalThis.atob(raw);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  // Test/runtime fallback (Node).
  if (typeof Buffer !== "undefined") {
    return Uint8Array.from(Buffer.from(raw, "base64"));
  }
  throw new Error("Base64 decoding is not available in this runtime");
}

function pemToDerBytes(pem: string): Uint8Array {
  const raw = String(pem)
    .trim()
    .replace(/-----BEGIN PUBLIC KEY-----/g, "")
    .replace(/-----END PUBLIC KEY-----/g, "")
    .replace(/\s+/g, "");
  if (!raw) throw new Error("Invalid public key PEM (empty)");
  return base64ToBytes(raw);
}

function isEd25519NotSupportedError(error: unknown): boolean {
  const name = typeof (error as any)?.name === "string" ? (error as any).name : "";
  const message = String((error as any)?.message ?? error);

  if (name === "NotSupportedError") return true;
  if (/does not support ed25519/i.test(message)) return true;
  if (/ed25519/i.test(message) && /(not supported|unsupported|does not support)/i.test(message)) return true;

  return false;
}

async function verifyEd25519SignatureViaWebCrypto(
  payloadBytes: Uint8Array,
  signatureBase64: string,
  publicKeyPem: string
): Promise<boolean> {
  const signature = base64ToBytes(signatureBase64);

  const subtle = globalThis.crypto?.subtle;
  if (!subtle || typeof subtle.importKey !== "function" || typeof subtle.verify !== "function") {
    throw new Error("WebCrypto (crypto.subtle.importKey()/verify()) is required to verify extension packages");
  }

  const key = await subtle.importKey("spki", pemToDerBytes(publicKeyPem), { name: "Ed25519" }, false, ["verify"]);
  const ok = await subtle.verify({ name: "Ed25519" }, key, signature, payloadBytes);
  return Boolean(ok);
}

async function verifyEd25519SignatureViaTauri(
  payloadBytes: Uint8Array,
  signatureBase64: string,
  publicKeyPem: string
): Promise<boolean> {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke;
  if (typeof invoke !== "function") {
    throw new Error(
      "This environment's WebCrypto implementation does not support Ed25519 and Tauri invoke() is not available."
    );
  }

  const ok = await invoke("verify_ed25519_signature", {
    payload: Array.from(payloadBytes),
    signature_base64: signatureBase64,
    public_key_pem: publicKeyPem
  });
  return Boolean(ok);
}

async function verifyEd25519SignatureDesktop(
  payloadBytes: Uint8Array,
  signatureBase64: string,
  publicKeyPem: string
): Promise<boolean> {
  try {
    return await verifyEd25519SignatureViaWebCrypto(payloadBytes, signatureBase64, publicKeyPem);
  } catch (error) {
    if (isEd25519NotSupportedError(error)) {
      return await verifyEd25519SignatureViaTauri(payloadBytes, signatureBase64, publicKeyPem);
    }

    const message = String((error as any)?.message ?? error);
    throw new Error(`Failed to verify extension signature via WebCrypto: ${message}`);
  }
}

export async function verifyExtensionPackageV2Desktop(
  packageBytes: Uint8Array,
  publicKeyPem: string
): Promise<VerifiedExtensionPackageV2> {
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
  const checksumEntries = new Map<string, { sha256: string; size: number }>();
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
    const sha = typeof (entry as any).sha256 === "string" ? String((entry as any).sha256).toLowerCase() : null;
    if (!sha || !/^[0-9a-f]{64}$/.test(sha)) {
      throw new Error(`checksums.json entry for ${rawPath} has invalid sha256`);
    }
    const size = (entry as any).size;
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
  const ok = await verifyEd25519SignatureDesktop(payloadBytes, signature.signatureBase64, publicKeyPem);
  if (!ok) throw new Error("Package signature verification failed");

  const fileRecords: Array<{ path: string; sha256: string; size: number }> = [];
  let unpackedSize = 0;
  let readme = "";

  const packageJsonBytes = files.get("package.json");
  if (!packageJsonBytes) {
    throw new Error("Invalid extension package: missing files/package.json");
  }
  let packageJson: any = null;
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
