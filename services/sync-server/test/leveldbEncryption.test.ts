import assert from "node:assert/strict";
import test from "node:test";

import {
  LEVELDB_ENCRYPTION_MAGIC,
  decryptLeveldbValue,
  encryptLeveldbValue,
} from "../src/leveldbEncryption.js";

test("encrypt/decrypt round-trip", () => {
  const key = Buffer.alloc(32, 7);
  const plaintext = Buffer.from("hello world");

  const encrypted = encryptLeveldbValue(plaintext, key);

  assert.ok(
    encrypted
      .subarray(0, LEVELDB_ENCRYPTION_MAGIC.byteLength)
      .equals(LEVELDB_ENCRYPTION_MAGIC)
  );
  assert.ok(!encrypted.equals(plaintext));

  const decrypted = decryptLeveldbValue(encrypted, key, true);
  assert.ok(decrypted.equals(plaintext));
});

test("strict mode rejects plaintext reads", () => {
  const key = Buffer.alloc(32, 1);
  assert.throws(
    () => decryptLeveldbValue(Buffer.from("plaintext"), key, true),
    /not encrypted/i
  );
});

