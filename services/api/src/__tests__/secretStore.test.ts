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

function encryptV1(key: Buffer, plaintext: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  const packed = Buffer.concat([iv, tag, ciphertext]).toString("base64");
  return `v1:${packed}`;
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

  it("fails decryption on AAD mismatch", () => {
    const key = crypto.randomBytes(32);
    const keyring: SecretStoreKeyring = { currentKeyId: "k1", keys: { k1: key } };
    const encrypted = encryptSecretValue(keyring, "secret-a", "hello");
    expect(() => decryptSecretValue(keyring, "secret-b", encrypted)).toThrow();
  });
});

describe("secrets rotation (integration)", () => {
  async function createDb(): Promise<Pool> {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    const db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });
    return db;
  }

  it("re-encrypts legacy secrets into the latest encoding", async () => {
    const db = await createDb();
    try {
      const oldKey = crypto.randomBytes(32);
      const newKey = crypto.randomBytes(32);
      const keyring: SecretStoreKeyring = { currentKeyId: "new", keys: { old: oldKey, new: newKey } };

      const alphaV1 = encryptV1(oldKey, "alpha-value");
      const betaOld = encryptSecretValue({ currentKeyId: "old", keys: keyring.keys }, "beta", "beta-value");
      const gammaCurrent = encryptSecretValue(keyring, "gamma", "gamma-value");

      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["alpha", alphaV1]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["beta", betaOld]);
      await db.query("INSERT INTO secrets (name, encrypted_value) VALUES ($1, $2)", ["gamma", gammaCurrent]);

      const result = await runSecretsRotation(db, keyring, { batchSize: 2 });
      expect(result).toEqual({ scanned: 3, rotated: 2, failed: 0 });

      const gammaRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", ["gamma"]);
      expect(String(gammaRow.rows[0].encrypted_value)).toBe(gammaCurrent);

      expect(await getSecret(db, keyring, "alpha")).toBe("alpha-value");
      expect(await getSecret(db, keyring, "beta")).toBe("beta-value");
      expect(await getSecret(db, keyring, "gamma")).toBe("gamma-value");
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

  it("lists and deletes secrets", async () => {
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

      expect(await listSecrets(db)).toEqual(["alpha", "pre%fix", "pre_fix", "prefix"]);
      expect(await listSecrets(db, "pre%")).toEqual(["pre%fix"]);
      expect(await listSecrets(db, "pre_")).toEqual(["pre_fix"]);

      await deleteSecret(db, "alpha");
      expect(await listSecrets(db)).toEqual(["pre%fix", "pre_fix", "prefix"]);
    } finally {
      await db.end();
    }
  });
});
