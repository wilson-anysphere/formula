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

export type SecretListEntry = {
  name: string;
  createdAt: Date;
  updatedAt: Date;
};

export function deriveSecretStoreKey(secret: string): Buffer {
  // AES-256 requires a 32-byte key. Hash the configured secret into a fixed-length key.
  return crypto.createHash("sha256").update(secret, "utf8").digest();
}

const SECRET_STORE_AAD_CONTEXT_V2 = "formula-secret-store";
const SECRET_STORE_AAD_CONTEXT_V2_LEGACY = "secret";

function secretAad(name: string, context: string): Buffer {
  // Cryptographic binding between ciphertext and the intended secret name.
  // Including a fixed context string ensures domain separation across other
  // AES-GCM uses.
  return crypto.createHash("sha256").update(`${context}:${name}`, "utf8").digest();
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
  cipher.setAAD(secretAad(name, SECRET_STORE_AAD_CONTEXT_V2));
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return packEncrypted(iv, tag, ciphertext);
}

function decryptV2WithContext(key: Buffer, name: string, packed: string, context: string): string {
  const { iv, tag, ciphertext } = unpackEncrypted(packed);
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAAD(secretAad(name, context));
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8");
}

function decryptV2(key: Buffer, name: string, packed: string): string {
  // The earliest v2 implementation used `secret:${name}` as the AAD context. We
  // now use `formula-secret-store:${name}` but keep decrypt compatibility so
  // existing rows continue to work.
  const aadContexts = [SECRET_STORE_AAD_CONTEXT_V2, SECRET_STORE_AAD_CONTEXT_V2_LEGACY];

  let lastErr: unknown;
  for (const context of aadContexts) {
    try {
      return decryptV2WithContext(key, name, packed, context);
    } catch (err) {
      lastErr = err;
    }
  }

  throw lastErr instanceof Error ? lastErr : new Error("Failed to decrypt secret");
}

/**
 * Decrypt a v2 secret using the *current* AAD context only.
 *
 * This is primarily intended for rotation tooling so we can detect whether a v2
 * row is still using the legacy AAD context (`secret:${name}`) even when the
 * key id matches the current key.
 */
export function decryptSecretValueV2CurrentAad(keyring: SecretStoreKeyring, name: string, value: string): string {
  const info = getSecretEncodingInfo(value);
  if (info.version !== "v2") throw new Error("Not a v2 secret");
  const firstColon = value.indexOf(":", 3);
  if (firstColon === -1) throw new Error("Invalid secret encoding");
  const packed = value.slice(firstColon + 1);

  const key = keyring.keys[info.keyId];
  if (!key) throw new Error(`Secret store key not found for keyId=${info.keyId}`);
  if (key.byteLength !== 32) throw new Error(`Secret store key ${info.keyId} must be 32 bytes (got ${key.byteLength})`);
  return decryptV2WithContext(key, name, packed, SECRET_STORE_AAD_CONTEXT_V2);
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

export async function secretExists(db: Queryable, name: string): Promise<boolean> {
  const result = await db.query("SELECT 1 FROM secrets WHERE name = $1", [name]);
  return (result.rowCount ?? 0) > 0;
}

export async function deleteSecret(db: Queryable, name: string): Promise<void> {
  await db.query("DELETE FROM secrets WHERE name = $1", [name]);
}

export async function listSecrets(db: Queryable, options: { prefix?: string } = {}): Promise<SecretListEntry[]> {
  const prefix = typeof options.prefix === "string" ? options.prefix : "";
  if (!prefix) {
    const result = await db.query("SELECT name, created_at, updated_at FROM secrets ORDER BY name ASC");
    return (result.rows as Array<{ name: string; created_at: Date; updated_at: Date }>).map((row) => ({
      name: String(row.name),
      createdAt: new Date(row.created_at),
      updatedAt: new Date(row.updated_at)
    }));
  }

  // Use a literal prefix match rather than `LIKE`, so prefixes containing `%` or
  // `_` are treated as plain characters. This also avoids `LIKE ... ESCAPE`,
  // which pg-mem does not fully support. We compute the prefix length in
  // JavaScript to avoid relying on SQL `length()`/`char_length()` (not
  // implemented by pg-mem).
  const prefixLength = Array.from(prefix).length;
  const result = await db.query(
    `
      SELECT name, created_at, updated_at
      FROM secrets
      WHERE substring(name, 1, $2) = $1
      ORDER BY name ASC
    `,
    [prefix, prefixLength]
  );
  return (result.rows as Array<{ name: string; created_at: Date; updated_at: Date }>).map((row) => ({
    name: String(row.name),
    createdAt: new Date(row.created_at),
    updatedAt: new Date(row.updated_at)
  }));
}
