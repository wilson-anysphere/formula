import crypto from "node:crypto";
import type { Pool, PoolClient } from "pg";

type Queryable = Pick<Pool, "query"> | Pick<PoolClient, "query">;

export type SecretStoreKeyring = {
  /**
   * Key id to use for new writes.
   */
  currentKeyId: string;
  /**
   * Map of key id -> raw 32-byte AES-256 key.
   */
  keys: Record<string, Buffer>;
};

export type SecretEncodingInfo =
  | { version: "v1" }
  | {
      version: "v2";
      keyId: string;
    };

export function deriveSecretStoreKey(secret: string): Buffer {
  // AES-256 requires a 32-byte key. Hash the configured secret into a fixed-length key.
  return crypto.createHash("sha256").update(secret, "utf8").digest();
}

function secretAad(name: string): Buffer {
  // Cryptographic binding between ciphertext and the intended secret name.
  return crypto.createHash("sha256").update(`secret:${name}`, "utf8").digest();
}

function packEncrypted(iv: Buffer, tag: Buffer, ciphertext: Buffer): string {
  return Buffer.concat([iv, tag, ciphertext]).toString("base64");
}

function unpackEncrypted(packed: string): { iv: Buffer; tag: Buffer; ciphertext: Buffer } {
  const raw = Buffer.from(packed, "base64");
  if (raw.byteLength < 12 + 16) throw new Error("Invalid secret encoding");
  const iv = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const ciphertext = raw.subarray(28);
  return { iv, tag, ciphertext };
}

function decryptV1(key: Buffer, packed: string): string {
  const { iv, tag, ciphertext } = unpackEncrypted(packed);
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8");
}

function encryptV2(key: Buffer, name: string, plaintext: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  cipher.setAAD(secretAad(name));
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return packEncrypted(iv, tag, ciphertext);
}

function decryptV2(key: Buffer, name: string, packed: string): string {
  const { iv, tag, ciphertext } = unpackEncrypted(packed);
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAAD(secretAad(name));
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8");
}

export function getSecretEncodingInfo(value: string): SecretEncodingInfo {
  if (value.startsWith("v1:")) return { version: "v1" };
  if (!value.startsWith("v2:")) throw new Error("Unsupported secret encoding");

  const firstColon = value.indexOf(":", 3);
  if (firstColon === -1) throw new Error("Invalid secret encoding");
  const keyId = value.slice(3, firstColon);
  if (!keyId) throw new Error("Invalid secret encoding");
  return { version: "v2", keyId };
}

export function encryptSecretValue(keyring: SecretStoreKeyring, name: string, plaintext: string): string {
  const keyId = keyring.currentKeyId;
  if (!keyId) throw new Error("Secret store currentKeyId must be set");
  if (keyId.includes(":")) throw new Error("Secret store currentKeyId must not contain ':'");
  const key = keyring.keys[keyId];
  if (!key) throw new Error(`Secret store key not found for currentKeyId=${keyId}`);
  if (key.byteLength !== 32) throw new Error(`Secret store key ${keyId} must be 32 bytes (got ${key.byteLength})`);
  const packed = encryptV2(key, name, plaintext);
  return `v2:${keyId}:${packed}`;
}

export function decryptSecretValue(keyring: SecretStoreKeyring, name: string, value: string): string {
  const info = getSecretEncodingInfo(value);
  if (info.version === "v2") {
    const firstColon = value.indexOf(":", 3);
    const packed = value.slice(firstColon + 1);
    const key = keyring.keys[info.keyId];
    if (!key) throw new Error(`Secret store key not found for keyId=${info.keyId}`);
    if (key.byteLength !== 32) throw new Error(`Secret store key ${info.keyId} must be 32 bytes (got ${key.byteLength})`);
    return decryptV2(key, name, packed);
  }

  // v1 secrets do not include key id or AAD. Best-effort: try all configured
  // keys until one successfully decrypts.
  const packed = value.slice(3);
  const entries = Object.entries(keyring.keys);
  if (entries.length === 0) throw new Error("No secret store keys configured");

  let lastErr: unknown;
  for (const [keyId, key] of entries) {
    try {
      if (key.byteLength !== 32) throw new Error(`Secret store key ${keyId} must be 32 bytes (got ${key.byteLength})`);
      return decryptV1(key, packed);
    } catch (err) {
      lastErr = err;
    }
  }

  throw lastErr instanceof Error ? lastErr : new Error("Failed to decrypt secret");
}

export async function putSecret(
  db: Queryable,
  keyring: SecretStoreKeyring,
  name: string,
  plaintext: string
): Promise<void> {
  const encrypted = encryptSecretValue(keyring, name, plaintext);
  await db.query(
    `
      INSERT INTO secrets (name, encrypted_value)
      VALUES ($1, $2)
      ON CONFLICT (name)
      DO UPDATE SET encrypted_value = EXCLUDED.encrypted_value, updated_at = now()
    `,
    [name, encrypted]
  );
}

export async function getSecret(
  db: Queryable,
  keyring: SecretStoreKeyring,
  name: string
): Promise<string | null> {
  const result = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [name]);
  if (result.rowCount !== 1) return null;
  const row = result.rows[0] as { encrypted_value: string };
  return decryptSecretValue(keyring, name, row.encrypted_value);
}

export async function deleteSecret(db: Queryable, name: string): Promise<void> {
  await db.query("DELETE FROM secrets WHERE name = $1", [name]);
}

export async function listSecrets(db: Queryable, prefix?: string): Promise<string[]> {
  if (!prefix) {
    const result = await db.query("SELECT name FROM secrets ORDER BY name ASC");
    return result.rows.map((row: any) => String(row.name));
  }

  // Avoid `LIKE ... ESCAPE` here: pg-mem does not fully support it and we want
  // prefix matching to treat `%`/`_` as literal characters. Instead, express the
  // prefix query as a range scan on the primary key.
  const lower = prefix;
  // U+10FFFF is the maximum valid Unicode code point; appending it provides an
  // exclusive upper bound that includes all strings starting with `prefix`.
  const upper = `${prefix}\u{10FFFF}`;
  const result = await db.query(
    "SELECT name FROM secrets WHERE name >= $1 AND name < $2 ORDER BY name ASC",
    [lower, upper]
  );
  return result.rows.map((row: any) => String(row.name));
}
