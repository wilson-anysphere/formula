import type { VerifiedExtensionPackageV2 } from "../../../../../shared/extension-package/v2-browser.mjs";
// `shared/` is a CommonJS workspace package, but the desktop webview/dev runtime executes ESM.
// Import the ESM wrapper so Vite can serve it without relying on CommonJS transforms (which
// are only applied during production builds).
import v2Core from "../../../../../shared/extension-package/core/v2-core.mjs";

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
// NOTE: Keep these limits in sync with:
// - `shared/extension-package/v2-browser.mjs`
// - `apps/desktop/src-tauri/src/ed25519_verifier.rs`
const MAX_SIGNATURE_PAYLOAD_BYTES = 5 * 1024 * 1024; // 5MB
const MAX_PUBLIC_KEY_PEM_BYTES = 64 * 1024; // 64KB
const MAX_SIGNATURE_BASE64_BYTES = 1024;

function bytesToHex(bytes: Uint8Array): string {
  let out = "";
  for (const b of bytes) out += b.toString(16).padStart(2, "0");
  return out;
}

function normalizeBytesForCrypto(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  // `crypto.subtle.*` expects BufferSources backed by `ArrayBuffer`. TypeScript models `Uint8Array`
  // as potentially backed by a `SharedArrayBuffer` (`ArrayBufferLike`), so normalize to an
  // `ArrayBuffer`-backed view for type safety (and to ensure downstream code doesn't need casts).
  return bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle || typeof subtle.digest !== "function") {
    throw new Error("WebCrypto subtle.digest is required to verify extension packages");
  }

  const normalized = normalizeBytesForCrypto(bytes);

  const digest = await subtle.digest("SHA-256", normalized);
  return bytesToHex(new Uint8Array(digest));
}

function base64ToBytes(base64: string): Uint8Array<ArrayBuffer> {
  const raw = String(base64);
  // Test/runtime fallback (Node).
  if (typeof Buffer !== "undefined") {
    const buf = Buffer.from(raw, "base64");
    const sliced = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    return new Uint8Array(sliced);
  }
  if (typeof globalThis.atob === "function") {
    const bin = globalThis.atob(raw);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
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
  const normalizedSignature = normalizeBytesForCrypto(signature);
  const normalizedPayload = normalizeBytesForCrypto(payloadBytes);
  const normalizedPublicKeyDer = normalizeBytesForCrypto(pemToDerBytes(publicKeyPem));

  const subtle = globalThis.crypto?.subtle;
  if (!subtle || typeof subtle.importKey !== "function" || typeof subtle.verify !== "function") {
    throw new Error("WebCrypto (crypto.subtle.importKey()/verify()) is required to verify extension packages");
  }

  const key = await subtle.importKey("spki", normalizedPublicKeyDer, { name: "Ed25519" }, false, ["verify"]);
  const ok = await subtle.verify({ name: "Ed25519" }, key, normalizedSignature, normalizedPayload);
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

  try {
    const ok = await invoke("verify_ed25519_signature", {
      payload: Array.from(payloadBytes),
      signature_base64: signatureBase64,
      public_key_pem: publicKeyPem
    });
    return Boolean(ok);
  } catch (error) {
    const message = String((error as any)?.message ?? error);
    throw new Error(`Failed to verify extension signature via Tauri IPC: ${message}`);
  }
}

async function verifyEd25519SignatureDesktop(
  payloadBytes: Uint8Array,
  signatureBase64: string,
  publicKeyPem: string
): Promise<boolean> {
  if (payloadBytes.length > MAX_SIGNATURE_PAYLOAD_BYTES) {
    throw new Error(`Signature payload is too large (max ${MAX_SIGNATURE_PAYLOAD_BYTES} bytes)`);
  }
  if (String(publicKeyPem).length > MAX_PUBLIC_KEY_PEM_BYTES) {
    throw new Error(`Public key PEM is too large (max ${MAX_PUBLIC_KEY_PEM_BYTES} bytes)`);
  }
  if (String(signatureBase64).length > MAX_SIGNATURE_BASE64_BYTES) {
    throw new Error(`Signature base64 is too large (max ${MAX_SIGNATURE_BASE64_BYTES} bytes)`);
  }

  try {
    return await verifyEd25519SignatureViaWebCrypto(payloadBytes, signatureBase64, publicKeyPem);
  } catch (error) {
    if (isEd25519NotSupportedError(error)) {
      const message = String((error as any)?.message ?? error);
      try {
        return await verifyEd25519SignatureViaTauri(payloadBytes, signatureBase64, publicKeyPem);
      } catch (invokeError) {
        const invokeMsg = String((invokeError as any)?.message ?? invokeError);
        throw new Error(
          "This environment's WebCrypto implementation does not support Ed25519, so extension packages cannot be verified. " +
            "If you are running Formula Desktop, ensure the native verifier is available. " +
            `Original error: ${message}. Tauri error: ${invokeMsg}`
        );
      }
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
