/**
 * Helpers for dealing with the key/signature formats used by Tauri's updater.
 *
 * Tauri uses minisign-compatible Ed25519 keys. The values we see in CI are typically:
 * - `plugins.updater.pubkey` in `tauri.conf.json`: base64 of a minisign public key file:
 *     untrusted comment: minisign public key: <KEYID>
 *     <base64 payload>
 *   where the payload decodes to:
 *     b"Ed" + keyid_le(8) + ed25519_pubkey(32)
 *
 * - `latest.json.sig`: either a raw Ed25519 signature (64 bytes, base64 encoded) or a minisign
 *   signature payload/file where the payload decodes to:
 *     b"Ed" + keyid_le(8) + ed25519_signature(64)
 */

/**
 * Normalizes base64/base64url and adds missing padding.
 *
 * Node's `Buffer.from(..., "base64")` is permissive (it will happily ignore invalid characters),
 * which can mask formatting problems. We validate the alphabet/padding to produce actionable errors.
 *
 * @param {string} label
 * @param {string} value
 * @returns {string}
 */
function normalizeBase64(label, value) {
  const normalized = value.trim().replace(/\s+/g, "");
  if (normalized.length === 0) {
    throw new Error(`${label} is empty.`);
  }

  // Support both standard base64 and base64url.
  let base64 = normalized.replace(/-/g, "+").replace(/_/g, "/");

  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) {
    throw new Error(`${label} is not valid base64.`);
  }

  // Allow unpadded base64 by adding the required '=' chars.
  const mod = base64.length % 4;
  if (mod === 1) {
    throw new Error(`${label} is not valid base64 (invalid length).`);
  }
  if (mod !== 0) {
    base64 += "=".repeat(4 - mod);
  }

  return base64;
}

/**
 * @param {string} label
 * @param {string} value
 * @returns {Buffer}
 */
function decodeBase64(label, value) {
  const normalized = normalizeBase64(label, value);
  return Buffer.from(normalized, "base64");
}

/**
 * @param {string} text
 * @returns {string[]}
 */
function nonEmptyLines(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

/**
 * @param {string} line
 * @returns {string | null}
 */
function extractKeyIdFromPublicKeyComment(line) {
  const match = line.match(/minisign public key:\s*([0-9a-fA-F]{16})\b/);
  return match ? match[1].toUpperCase() : null;
}

/**
 * Minisign encodes key ids as little-endian bytes, while the comment line prints the key id as
 * uppercase hex in big-endian order.
 *
 * @param {Uint8Array} keyIdLe8
 * @returns {string}
 */
function formatKeyId(keyIdLe8) {
  if (keyIdLe8.length !== 8) return Buffer.from(keyIdLe8).toString("hex").toUpperCase();
  return Buffer.from(keyIdLe8).reverse().toString("hex").toUpperCase();
}

/**
 * Parses a minisign Ed25519 public key payload (the `RW...` line in a minisign public key file).
 *
 * Payload binary format (42 bytes):
 *   b"Ed" + keyid_le(8) + ed25519_pubkey(32)
 *
 * @param {string} payloadBase64
 * @returns {{ publicKeyBytes: Buffer, keyId: string }}
 */
export function parseMinisignPublicKeyPayload(payloadBase64) {
  const bytes = decodeBase64("minisign public key payload", payloadBase64);
  if (bytes.length !== 42) {
    throw new Error(
      `minisign public key payload decoded to ${bytes.length} bytes (expected 42: b\"Ed\" + keyid(8) + pubkey(32)).`,
    );
  }
  if (bytes[0] !== 0x45 || bytes[1] !== 0x64) {
    throw new Error(
      `minisign public key payload has invalid prefix (expected 0x45 0x64 / \"Ed\", got 0x${bytes[0]?.toString(16)} 0x${bytes[1]?.toString(16)}).`,
    );
  }

  const keyIdLe = bytes.subarray(2, 10);
  const pubkey = bytes.subarray(10);
  if (pubkey.length !== 32) {
    throw new Error(`minisign public key payload has ${pubkey.length} pubkey bytes (expected 32).`);
  }

  return { publicKeyBytes: Buffer.from(pubkey), keyId: formatKeyId(keyIdLe) };
}

/**
 * Parses the `plugins.updater.pubkey` value from `tauri.conf.json`.
 *
 * - If the base64 decodes to 32 bytes, it's treated as a raw Ed25519 public key.
 * - If the base64 decodes to 42 bytes and starts with `b"Ed"`, it's treated as a raw minisign
 *   public key payload (`b"Ed" + keyid_le(8) + pubkey(32)`).
 * - Otherwise it is treated as base64 of a minisign public key file (2-line text block).
 *
 * @param {string} pubkeyBase64
 * @returns {{ publicKeyBytes: Buffer, keyId: string | null, format: "raw" | "minisign" }}
 */
export function parseTauriUpdaterPubkey(pubkeyBase64) {
  const decoded = decodeBase64("updater pubkey", pubkeyBase64);
  if (decoded.length === 32) {
    return { publicKeyBytes: decoded, keyId: null, format: "raw" };
  }

  if (decoded.length === 42 && decoded[0] === 0x45 && decoded[1] === 0x64) {
    const keyIdLe = decoded.subarray(2, 10);
    const pubkey = decoded.subarray(10);
    if (pubkey.length !== 32) {
      throw new Error(`minisign public key payload has ${pubkey.length} pubkey bytes (expected 32).`);
    }
    return { publicKeyBytes: Buffer.from(pubkey), keyId: formatKeyId(keyIdLe), format: "minisign" };
  }

  const text = decoded.toString("utf8").trim();
  if (!text) {
    throw new Error(
      `updater pubkey decoded to ${decoded.length} bytes, but the result is not valid minisign text.`,
    );
  }

  const lines = nonEmptyLines(text);
  if (lines.length === 1) {
    // Some setups may store only the minisign payload line (without the comment).
    const { publicKeyBytes, keyId } = parseMinisignPublicKeyPayload(lines[0]);
    return { publicKeyBytes, keyId, format: "minisign" };
  }

  if (lines.length < 2) {
    throw new Error(
      `updater pubkey decoded to minisign text with ${lines.length} line(s); expected at least 2 (comment + payload).`,
    );
  }

  const commentLine = lines[0];
  const payloadLine = lines[1];
  const { publicKeyBytes, keyId } = parseMinisignPublicKeyPayload(payloadLine);

  const commentKeyId = extractKeyIdFromPublicKeyComment(commentLine);
  if (commentKeyId && commentKeyId !== keyId) {
    throw new Error(
      `minisign public key comment key id (${commentKeyId}) does not match payload key id (${keyId}).`,
    );
  }

  return { publicKeyBytes, keyId, format: "minisign" };
}

/**
 * Parses an updater signature file (e.g. `latest.json.sig`) into a raw Ed25519 signature.
 *
 * Supported inputs:
 * - raw base64 Ed25519 signature (64 bytes)
 * - minisign signature payload (base64 of 74 bytes: b"Ed" + keyid_le(8) + signature(64))
 * - minisign signature file (text): "untrusted comment: ...\n<payload>\n[...]"
 *
 * @param {string} signatureText
 * @param {string} [label]
 * @returns {{ signatureBytes: Buffer, keyId: string | null, format: "raw" | "minisign" }}
 */
export function parseTauriUpdaterSignature(signatureText, label = "signature") {
  const trimmed = signatureText.trim();
  if (!trimmed) {
    throw new Error(`${label} is empty.`);
  }

  const lines = nonEmptyLines(trimmed);
  /** @type {string} */
  let base64Line;

  if (lines.length >= 2 && lines[0].startsWith("untrusted comment:")) {
    // minisign signature file (2 or 4 lines); the payload is always the second line.
    base64Line = lines[1];
  } else if (lines.length === 1) {
    base64Line = lines[0];
  } else {
    // If this doesn't look like minisign (missing comment line), attempt to find a base64-like line.
    const maybe = lines.find((line) => {
      try {
        normalizeBase64(label, line);
        return true;
      } catch {
        return false;
      }
    });
    if (!maybe) {
      const preview = lines[0] ? JSON.stringify(lines[0].slice(0, 80)) : "(empty)";
      throw new Error(
        `${label} is not recognized as base64 or minisign signature text (first line: ${preview}).`,
      );
    }
    base64Line = maybe;
  }

  const bytes = decodeBase64(label, base64Line);

  if (bytes.length === 64) {
    return { signatureBytes: bytes, keyId: null, format: "raw" };
  }

  if (bytes.length === 74) {
    if (bytes[0] !== 0x45 || bytes[1] !== 0x64) {
      throw new Error(
        `${label} minisign payload has invalid prefix (expected 0x45 0x64 / \"Ed\", got 0x${bytes[0]?.toString(16)} 0x${bytes[1]?.toString(16)}).`,
      );
    }
    const keyIdLe = bytes.subarray(2, 10);
    const signature = bytes.subarray(10);
    if (signature.length !== 64) {
      throw new Error(
        `${label} minisign payload has ${signature.length} signature bytes (expected 64).`,
      );
    }
    return { signatureBytes: Buffer.from(signature), keyId: formatKeyId(keyIdLe), format: "minisign" };
  }

  throw new Error(
    `${label} decoded to ${bytes.length} bytes; expected either 64 (raw Ed25519 signature) or 74 (minisign payload: b\"Ed\" + keyid(8) + sig(64)).`,
  );
}
