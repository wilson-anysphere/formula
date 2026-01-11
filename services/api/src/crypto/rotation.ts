import crypto from "node:crypto";
import type { Pool, PoolClient } from "pg";
import { withTransaction } from "../db/tx";
import { Keyring } from "./keyring";

export type KeyRotationResult = {
  orgsRotated: number;
  documentVersionDeksRewrapped: number;
};

const DAY_MS = 24 * 60 * 60 * 1000;

function needsRotation({
  now,
  rotatedAt,
  rotationDays
}: {
  now: Date;
  rotatedAt: Date;
  rotationDays: number;
}): boolean {
  return now.getTime() - rotatedAt.getTime() >= rotationDays * DAY_MS;
}

async function rewrapDocumentVersionDeks({
  client,
  keyring,
  orgId,
  targetKmsProvider,
  targetKmsKeyId
}: {
  client: PoolClient;
  keyring: Keyring;
  orgId: string;
  targetKmsProvider: string;
  targetKmsKeyId: string;
}): Promise<number> {
  const rows = await client.query(
    `
      SELECT v.id, v.data_encrypted_dek, v.data_kms_provider, v.data_kms_key_id
      FROM document_versions v
      JOIN documents d ON d.id = v.document_id
      WHERE d.org_id = $1
        AND v.data_encrypted_dek IS NOT NULL
    `,
    [orgId]
  );

  let updated = 0;
  for (const row of rows.rows as any[]) {
    const encryptedDek = Buffer.from(String(row.data_encrypted_dek), "base64");
    const sourceProvider = String(row.data_kms_provider ?? targetKmsProvider);
    const sourceKmsKeyId = String(row.data_kms_key_id ?? targetKmsKeyId);

    const sourceKms = keyring.get(sourceProvider);
    const plaintextDek = await sourceKms.decryptKey({
      encryptedDek,
      orgId,
      kmsKeyId: sourceKmsKeyId
    });

    const targetKms = keyring.get(targetKmsProvider);
    const rewrapped = await targetKms.encryptKey({
      plaintextDek,
      orgId,
      keyId: targetKmsKeyId
    });

    await client.query(
      `
        UPDATE document_versions
        SET data_encrypted_dek = $2,
            data_kms_provider = $3,
            data_kms_key_id = $4
        WHERE id = $1
      `,
      [String(row.id), rewrapped.encryptedDek.toString("base64"), targetKmsProvider, rewrapped.kmsKeyId]
    );
    updated += 1;
  }

  return updated;
}

/**
 * Rotate each org's configured KMS key id (org_settings.kms_key_id) when due, and
 * re-wrap stored DEKs to the latest key id (ciphertext remains untouched).
 */
export async function runKeyRotation(
  pool: Pool,
  keyring: Keyring,
  { now = new Date(), orgId }: { now?: Date; orgId?: string } = {}
): Promise<KeyRotationResult> {
  const orgs = await pool.query(
    `
      SELECT
        org_id,
        cloud_encryption_at_rest,
        kms_provider,
        kms_key_id,
        key_rotation_days,
        kms_key_rotated_at
      FROM org_settings
      ${orgId ? "WHERE org_id = $1" : ""}
    `,
    orgId ? [orgId] : []
  );

  let orgsRotated = 0;
  let documentVersionDeksRewrapped = 0;

  for (const row of orgs.rows as any[]) {
    const cloudEncryptionAtRest = Boolean(row.cloud_encryption_at_rest);
    if (!cloudEncryptionAtRest) continue;

    const rotationDays = Number(row.key_rotation_days ?? 90);
    const rotatedAtRaw = row.kms_key_rotated_at ? new Date(row.kms_key_rotated_at as string) : new Date(0);
    if (!needsRotation({ now, rotatedAt: rotatedAtRaw, rotationDays })) continue;

    const orgIdValue = String(row.org_id);
    const kmsProvider = String(row.kms_provider ?? "local");
    const currentKmsKeyId = row.kms_key_id == null ? null : String(row.kms_key_id);

    const nextKmsKeyId =
      kmsProvider === "local"
        ? `local-${crypto.randomUUID()}`
        : currentKmsKeyId ?? (() => {
            throw new Error(`kms_key_id is required for kms_provider=${kmsProvider}`);
          })();

    await withTransaction(pool, async (client) => {
      await client.query(
        `
          UPDATE org_settings
          SET kms_key_id = $2,
              kms_key_rotated_at = $3,
              updated_at = now()
          WHERE org_id = $1
        `,
        [orgIdValue, nextKmsKeyId, now]
      );

      documentVersionDeksRewrapped += await rewrapDocumentVersionDeks({
        client,
        keyring,
        orgId: orgIdValue,
        targetKmsProvider: kmsProvider,
        targetKmsKeyId: nextKmsKeyId
      });
    });

    orgsRotated += 1;
  }

  return { orgsRotated, documentVersionDeksRewrapped };
}
