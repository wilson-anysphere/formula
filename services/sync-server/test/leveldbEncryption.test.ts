import assert from "node:assert/strict";
import test from "node:test";

import levelMem from "level-mem";

import { KeyRing } from "../../../packages/security/crypto/keyring.js";
import {
  DEFAULT_LEVELDB_VALUE_MAGIC,
  RAW_VALUE_ENCODING,
  createEncryptedLevelAdapter,
} from "../src/leveldbEncryption.js";

function createTestKeyRing(): KeyRing {
  return KeyRing.fromJSON({
    currentVersion: 1,
    keys: {
      // 32 bytes of deterministic test key material.
      "1": Buffer.alloc(32, 7).toString("base64"),
    },
  });
}

test("createEncryptedLevelAdapter encrypts values via valueEncoding wrapper", async (t) => {
  const keyRing = createTestKeyRing();
  const encryptedLevel = createEncryptedLevelAdapter({
    keyRing,
    strict: true,
    magic: DEFAULT_LEVELDB_VALUE_MAGIC,
  })(levelMem as any);

  const db = encryptedLevel(`mem-${Date.now()}-${Math.random().toString(16).slice(2)}`, {
    valueEncoding: RAW_VALUE_ENCODING,
  });
  t.after(async () => {
    await db.close();
  });

  const plaintext = Buffer.from("hello world");
  await db.put("k", plaintext);

  const roundTrip = (await db.get("k")) as Buffer;
  assert.ok(Buffer.isBuffer(roundTrip));
  assert.ok(roundTrip.equals(plaintext));

  const rawStored = (await db.get("k", { valueEncoding: RAW_VALUE_ENCODING })) as Buffer;
  assert.ok(Buffer.isBuffer(rawStored));
  assert.ok(
    rawStored
      .subarray(0, DEFAULT_LEVELDB_VALUE_MAGIC.byteLength)
      .equals(DEFAULT_LEVELDB_VALUE_MAGIC)
  );
  assert.ok(!rawStored.equals(plaintext));
});

test("strict mode rejects legacy plaintext reads", async (t) => {
  const keyRing = createTestKeyRing();
  const encryptedLevel = createEncryptedLevelAdapter({
    keyRing,
    strict: true,
    magic: DEFAULT_LEVELDB_VALUE_MAGIC,
  })(levelMem as any);

  const db = encryptedLevel(`mem-${Date.now()}-${Math.random().toString(16).slice(2)}`, {
    valueEncoding: RAW_VALUE_ENCODING,
  });
  t.after(async () => {
    await db.close();
  });

  // Simulate a legacy plaintext record written without the encryption wrapper.
  await db.put("legacy", Buffer.from("plaintext"), { valueEncoding: RAW_VALUE_ENCODING });

  await assert.rejects(db.get("legacy"), /Encountered unencrypted LevelDB value/);
});
