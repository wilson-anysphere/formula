import assert from "node:assert/strict";
import test from "node:test";

import levelMem from "level-mem";
import * as Y from "yjs";

import { loadYLeveldbFromTarball } from "./y-leveldb-tarball.js";
import {
  DEFAULT_LEVELDB_VALUE_MAGIC,
  RAW_VALUE_ENCODING,
  createEncryptedLevelAdapter,
} from "../src/leveldbEncryption.js";
import { KeyRing } from "../../../packages/security/crypto/keyring.js";

function createTestKeyRing(): KeyRing {
  return KeyRing.fromJSON({
    currentVersion: 1,
    keys: {
      // 32 bytes of deterministic test key material.
      "1": Buffer.alloc(32, 9).toString("base64"),
    },
  });
}

test("LeveldbPersistence stores encrypted values via custom level adapter", async (t) => {
  const { LeveldbPersistence } = await loadYLeveldbFromTarball(t);

  const location = `mem-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const keyRing = createTestKeyRing();

  const encryptedLevel = createEncryptedLevelAdapter({
    keyRing,
    strict: true,
  })(levelMem as any);

  const ldb = new LeveldbPersistence(location, { level: encryptedLevel });
  t.after(async () => {
    await ldb.destroy();
  });

  const docName = "doc";
  const doc = new Y.Doc();
  doc.getText("t").insert(0, "hello");
  const update = Y.encodeStateAsUpdate(doc);

  await ldb.storeUpdate(docName, update);

  const ydoc = await ldb.getYDoc(docName);
  assert.equal(ydoc.getText("t").toString(), "hello");

  // y-leveldb doesn't expose the underlying DB instance, but it does expose its
  // internal transaction helper which captures the DB in a closure. Read the
  // stored bytes with an identity valueEncoding to bypass decryption.
  const stored = (await (ldb as any)._transact((db: any) =>
    db.get(["v1", docName, "update", 0], {
      valueEncoding: RAW_VALUE_ENCODING,
    })
  )) as Buffer | null;

  assert.ok(stored, "expected an encrypted value stored for the update key");
  assert.ok(Buffer.isBuffer(stored));
  assert.ok(
    stored
      .subarray(0, DEFAULT_LEVELDB_VALUE_MAGIC.byteLength)
      .equals(DEFAULT_LEVELDB_VALUE_MAGIC)
  );
  assert.ok(!stored.equals(Buffer.from(update)));
  assert.equal(
    stored.includes(Buffer.from("hello", "utf8")),
    false,
    "encrypted values should not contain plaintext UTF-8 substrings"
  );
});
