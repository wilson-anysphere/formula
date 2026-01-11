import crypto from "node:crypto";
import type { Pool, PoolClient } from "pg";

type Queryable = Pick<Pool, "query"> | Pick<PoolClient, "query">;

function deriveKey(secret: string): Buffer {
  // AES-256 requires a 32-byte key. Hash the configured secret into a fixed-length key.
  return crypto.createHash("sha256").update(secret, "utf8").digest();
}

function encryptValue(key: Buffer, plaintext: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();

  const packed = Buffer.concat([iv, tag, ciphertext]).toString("base64");
  return `v1:${packed}`;
}

function decryptValue(key: Buffer, value: string): string {
  if (!value.startsWith("v1:")) throw new Error("Unsupported secret encoding");
  const raw = Buffer.from(value.slice(3), "base64");
  if (raw.byteLength < 12 + 16) throw new Error("Invalid secret encoding");
  const iv = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const ciphertext = raw.subarray(28);

  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8");
}

export async function putSecret(
  db: Queryable,
  encryptionSecret: string,
  name: string,
  plaintext: string
): Promise<void> {
  const key = deriveKey(encryptionSecret);
  const encrypted = encryptValue(key, plaintext);
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
  encryptionSecret: string,
  name: string
): Promise<string | null> {
  const result = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [name]);
  if (result.rowCount !== 1) return null;
  const row = result.rows[0] as { encrypted_value: string };
  const key = deriveKey(encryptionSecret);
  return decryptValue(key, row.encrypted_value);
}

export async function deleteSecret(db: Queryable, name: string): Promise<void> {
  await db.query("DELETE FROM secrets WHERE name = $1", [name]);
}
