import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { runMigrations } from "../db/migrations";
import {
  deleteSecret,
  decryptSecretValue,
  encryptSecretValue,
  getSecret,
  listSecrets,
  type SecretStoreKeyring
} from "../secrets/secretStore";
import { runSecretsRotation } from "../secrets/rotation";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

async function createDb(): Promise<Pool> {
  const mem = newDb({ autoCreateForeignKeyIndices: true });
  const pgAdapter = mem.adapters.createPg();
  const db = new pgAdapter.Pool();
  await runMigrations(db, { migrationsDir: getMigrationsDir() });
  return db;
}

function encryptV1(key: Buffer, plaintext: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  const packed = Buffer.concat([iv, tag, ciphertext]).toString("base64");
  return `v1:${packed}`;
}

function encryptV2LegacyAad(key: Buffer, keyId: string, name: string, plaintext: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  // Legacy v2 AAD context was `secret:${name}` (SHA-256 digest).
  const aad = crypto.createHash("sha256").update(`secret:${name}`, "utf8").digest();
  cipher.setAAD(aad);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  const packed = Buffer.concat([iv, tag, ciphertext]).toString("base64");
  return `v2:${keyId}:${packed}`;
}

describe("secretStore", () => {
  it("roundtrips v2 secrets with AAD binding", () => {
    const key = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };

    const encrypted = encryptSecretValue(keyring, "my-secret", "hello");
    expect(encrypted.startsWith("v2:k1:")).toBe(true);

    const decrypted = decryptSecretValue(keyring, "my-secret", encrypted);
    expect(decrypted).toBe("hello");
  });

  it("supports key rotation (decrypt old after adding key; encrypt uses newest)", () => {
    const oldKey = crypto.randomBytes(32);
    const newKey = crypto.randomBytes(32);

    const oldKeyring: SecretStoreKeyring = { currentKeyId: "old", keys: { old: oldKey } };
    const rotatedKeyring: SecretStoreKeyring = { currentKeyId: "new", keys: { old: oldKey, new: newKey } };

    const encryptedOld = encryptSecretValue(oldKeyring, "rotating-secret", "value");
    expect(encryptedOld.startsWith("v2:old:")).toBe(true);
    expect(decryptSecretValue(rotatedKeyring, "rotating-secret", encryptedOld)).toBe("value");

    const encryptedNew = encryptSecretValue(rotatedKeyring, "rotating-secret", "value");
    expect(encryptedNew.startsWith("v2:new:")).toBe(true);
    expect(decryptSecretValue(rotatedKeyring, "rotating-secret", encryptedNew)).toBe("value");
  });

  it("rejects key ids containing ':' (would make encoding ambiguous)", () => {
    const key = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = { currentKeyId: "bad:key", keys: { "bad:key": key } };
    expect(() => encryptSecretValue(keyring, "my-secret", "hello")).toThrow();
  });

  it("decrypts legacy v1 secrets", () => {
    const currentKey = crypto.randomBytes(32);
    const oldKey = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = {
      currentKeyId: "current",
      keys: { current: currentKey, old: oldKey }
    };

    const encrypted = encryptV1(oldKey, "legacy");
    expect(decryptSecretValue(keyring, "ignored-name", encrypted)).toBe("legacy");
  });

  it("decrypts v2 secrets written with the legacy AAD context", () => {
    const key = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };

    const encrypted = encryptV2LegacyAad(key, "k1", "legacy-aad", "value");
    expect(decryptSecretValue(keyring, "legacy-aad", encrypted)).toBe("value");
  });

  it("fails decryption on AAD mismatch", () => {
    const key = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };
    const encrypted = encryptSecretValue(keyring, "secret-a", "hello");
    expect(() => decryptSecretValue(keyring, "secret-b", encrypted)).toThrow();
  });
});

describe("secrets rotation (integration)", () => {
  it("re-encrypts legacy secrets into the latest encoding", async () => {
    const db = await createDb();
    try {
      const oldKey = crypto.randomBytes(32);
      const newKey = crypto.randomBytes(32);
      const keyring: SecretStoreKeyring = { currentKeyId: "new", keys: { old: oldKey, new: newKey } };

      const alphaV1 = encryptV1(oldKey, "alpha-value");
      const betaOld = encryptSecretValue({ currentKeyId: "old", keys: keyring.keys }, "beta", "beta-value");
      const gammaCurrent = encryptSecretValue(keyring, "gamma", "gamma-value");
      const deltaLegacyAad = encryptV2LegacyAad(newKey, "new", "delta", "delta-value");

      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["alpha", alphaV1]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["beta", betaOld]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["gamma", gammaCurrent]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["delta", deltaLegacyAad]);

      const result = await runSecretsRotation(db, keyring, { batchSize: 2 });
      expect(result).toEqual({ scanned: 4, rotated: 3, failed: 0 });

      const gammaRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", ["gamma"]);
      expect(String(gammaRow.rows[0].encrypted_value)).toBe(gammaCurrent);

      expect(await getSecret(db, keyring, "alpha")).toBe("alpha-value");
      expect(await getSecret(db, keyring, "beta")).toBe("beta-value");
      expect(await getSecret(db, keyring, "gamma")).toBe("gamma-value");
      expect(await getSecret(db, keyring, "delta")).toBe("delta-value");
    } finally {
      await db.end();
    }
  });

  it("continues rotating when some secrets fail to decrypt", async () => {
    const db = await createDb();
    try {
      const key = crypto.randomBytes(32);
      const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };

      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["bad", "v2:missing:AAAA"]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["good", encryptV1(key, "ok")]);

      const result = await runSecretsRotation(db, keyring);
      expect(result).toEqual({ scanned: 2, rotated: 1, failed: 1 });
      expect(await getSecret(db, keyring, "good")).toBe("ok");
    } finally {
      await db.end();
    }
  });

  it("supports a prefix filter", async () => {
    const db = await createDb();
    try {
      const oldKey = crypto.randomBytes(32);
      const newKey = crypto.randomBytes(32);
      const keyring: SecretStoreKeyring = { currentKeyId: "new", keys: { old: oldKey, new: newKey } };

      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["pref:one", encryptV1(oldKey, "1")]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["other:one", encryptV1(oldKey, "2")]);

      const result = await runSecretsRotation(db, keyring, { prefix: "pref:" });
      expect(result.scanned).toBe(1);
      expect(result.rotated).toBe(1);
      expect(result.failed).toBe(0);

      expect(await getSecret(db, keyring, "pref:one")).toBe("1");
      expect(await getSecret(db, keyring, "other:one")).toBe("2");
    } finally {
      await db.end();
    }
  });

  it("lists and deletes secrets (prefix match is literal)", async () => {
    const db = await createDb();
    try {
      const key = crypto.randomBytes(32);
      const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };

      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", [
        "alpha",
        encryptSecretValue(keyring, "alpha", "a")
      ]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", [
        "pre%fix",
        encryptSecretValue(keyring, "pre%fix", "literal-percent")
      ]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", [
        "pre_fix",
        encryptSecretValue(keyring, "pre_fix", "literal-underscore")
      ]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", [
        "prefix",
        encryptSecretValue(keyring, "prefix", "prefix")
      ]);

      const all = await listSecrets(db);
      expect(all.map((row) => row.name)).toEqual(["alpha", "pre%fix", "pre_fix", "prefix"]);
      expect(all[0]!.createdAt).toBeInstanceOf(Date);
      expect(all[0]!.updatedAt).toBeInstanceOf(Date);

      expect((await listSecrets(db, { prefix: "pre%" })).map((row) => row.name)).toEqual(["pre%fix"]);
      expect((await listSecrets(db, { prefix: "pre_" })).map((row) => row.name)).toEqual(["pre_fix"]);

      await deleteSecret(db, "alpha");
      expect((await listSecrets(db)).map((row) => row.name)).toEqual(["pre%fix", "pre_fix", "prefix"]);
    } finally {
      await db.end();
    }
  });
});
