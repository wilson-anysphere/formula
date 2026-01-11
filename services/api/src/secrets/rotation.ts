import type { Pool } from "pg";
import {
  decryptSecretValue,
  encryptSecretValue,
  getSecretEncodingInfo,
  type SecretStoreKeyring
} from "./secretStore";

export type SecretsRotationResult = {
  scanned: number;
  rotated: number;
  failed: number;
};

type SecretRow = {
  name: string;
  encrypted_value: string;
};

/**
 * Re-encrypt all secrets not using the current key id (including legacy v1
 * secrets) into the latest encoding.
 *
 * The process is idempotent and safe to re-run; it is intended to be executed
 * while the API continues handling reads/writes.
 */
export async function runSecretsRotation(
  db: Pool,
  keyring: SecretStoreKeyring,
  { batchSize = 250 }: { batchSize?: number } = {}
): Promise<SecretsRotationResult> {
  let scanned = 0;
  let rotated = 0;
  let failed = 0;

  let lastName: string | null = null;

  // Cursor pagination over the primary key (name) avoids OFFSET scans and allows
  // the script to be restarted safely.
  while (true) {
    const res =
      lastName == null
        ? await db.query<SecretRow>(
            `
              SELECT name, encrypted_value
              FROM secrets
              ORDER BY name ASC
              LIMIT $1
            `,
            [batchSize]
          )
        : await db.query<SecretRow>(
            `
              SELECT name, encrypted_value
              FROM secrets
              WHERE name > $1
              ORDER BY name ASC
              LIMIT $2
            `,
            [lastName, batchSize]
          );

    if (res.rowCount === 0) break;

    for (const row of res.rows as any[]) {
      const name = String(row.name);
      const encryptedValue = String(row.encrypted_value);
      scanned += 1;

      lastName = name;

      let shouldRotate = true;
      try {
        const info = getSecretEncodingInfo(encryptedValue);
        if (info.version === "v2" && info.keyId === keyring.currentKeyId) {
          shouldRotate = false;
        }
      } catch {
        // Unknown/invalid encoding: we'll count it as failed below.
        shouldRotate = true;
      }

      if (!shouldRotate) continue;

      try {
        const plaintext = decryptSecretValue(keyring, name, encryptedValue);
        const reencrypted = encryptSecretValue(keyring, name, plaintext);

        const update = await db.query(
          `
            UPDATE secrets
            SET encrypted_value = $2,
                updated_at = now()
            WHERE name = $1 AND encrypted_value = $3
          `,
          [name, reencrypted, encryptedValue]
        );

        if (update.rowCount === 1) {
          rotated += 1;
        }
      } catch {
        failed += 1;
      }
    }
  }

  return { scanned, rotated, failed };
}

