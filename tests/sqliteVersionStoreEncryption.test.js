import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { InMemoryKeychainProvider } from "../packages/security/crypto/keychain/inMemoryKeychain.js";
import { isEncryptedFileBytes } from "../packages/security/crypto/encryptedFile.js";
import { SQLiteVersionStore } from "../packages/versioning/src/store/sqliteVersionStore.js";

test("SQLiteVersionStore: enable encryption -> ciphertext on disk; disable -> plaintext on disk", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "sqlite-store-encryption-"));
  const storePath = path.join(tmpDir, "versions.sqlite");

  const keychain = new InMemoryKeychainProvider();
  const service = "formula.test";
  const account = "sqlite-version-store-test";

  const store = new SQLiteVersionStore({
    filePath: storePath,
    encryption: {
      enabled: true,
      keychainProvider: keychain,
      keychainService: service,
      keychainAccount: account,
      aadContext: { scope: "test" }
    }
  });

  await store.saveVersion({
    id: "v1",
    kind: "snapshot",
    timestampMs: Date.now(),
    userId: null,
    userName: null,
    description: null,
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: Buffer.from([1, 2, 3, 4])
  });

  const encryptedBytes = await fs.readFile(storePath);
  assert.ok(isEncryptedFileBytes(encryptedBytes), "expected encrypted file format on disk");
  assert.ok(
    !encryptedBytes.toString("ascii").includes("SQLite format 3"),
    "expected ciphertext-only SQLite store"
  );

  const storeReload = new SQLiteVersionStore({
    filePath: storePath,
    encryption: {
      enabled: true,
      keychainProvider: keychain,
      keychainService: service,
      keychainAccount: account,
      aadContext: { scope: "test" }
    }
  });

  const loaded = await storeReload.getVersion("v1");
  assert.ok(loaded);
  assert.deepEqual(Buffer.from(loaded.snapshot), Buffer.from([1, 2, 3, 4]));

  await storeReload.disableEncryption();

  const plaintextBytes = await fs.readFile(storePath);
  assert.ok(
    plaintextBytes.subarray(0, 16).toString("ascii").startsWith("SQLite format 3"),
    "expected plaintext SQLite file after disabling encryption"
  );

  const secret = await keychain.getSecret({ service, account });
  assert.equal(secret, null, "expected keyring secret to be deleted on disable");

  const storePlain = new SQLiteVersionStore({ filePath: storePath });
  const loadedPlain = await storePlain.getVersion("v1");
  assert.ok(loadedPlain);
  assert.deepEqual(Buffer.from(loadedPlain.snapshot), Buffer.from([1, 2, 3, 4]));
});

