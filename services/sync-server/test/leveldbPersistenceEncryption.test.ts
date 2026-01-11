import assert from "node:assert/strict";
import test from "node:test";

import levelMem from "level-mem";
import { Y } from "./yjs-interop.ts";

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
  const secretText = "hello";
  doc.getText("t").insert(0, secretText);
  const update = Y.encodeStateAsUpdate(doc);

  await ldb.storeUpdate(docName, update);
  await ldb.setMeta(docName, "example", { hello: "world" });

  const ydoc = await ldb.getYDoc(docName);
  assert.equal(ydoc.getText("t").toString(), secretText);

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
    stored.includes(Buffer.from(secretText, "utf8")),
    false,
    "encrypted values should not contain plaintext UTF-8 substrings"
  );

  const stateVector = (await (ldb as any)._transact((db: any) =>
    db.get(["v1_sv", docName], {
      valueEncoding: RAW_VALUE_ENCODING,
    })
  )) as Buffer | null;
  assert.ok(stateVector, "expected an encrypted value stored for the state vector key");
  assert.ok(Buffer.isBuffer(stateVector));
  assert.ok(
    stateVector
      .subarray(0, DEFAULT_LEVELDB_VALUE_MAGIC.byteLength)
      .equals(DEFAULT_LEVELDB_VALUE_MAGIC)
  );

  const meta = (await (ldb as any)._transact((db: any) =>
    db.get(["v1", docName, "meta", "example"], {
      valueEncoding: RAW_VALUE_ENCODING,
    })
  )) as Buffer | null;
  assert.ok(meta, "expected an encrypted value stored for the meta key");
  assert.ok(Buffer.isBuffer(meta));
  assert.ok(
    meta.subarray(0, DEFAULT_LEVELDB_VALUE_MAGIC.byteLength).equals(DEFAULT_LEVELDB_VALUE_MAGIC)
  );
  assert.equal(
    meta.includes(Buffer.from("world", "utf8")),
    false,
    "encrypted meta should not contain plaintext UTF-8 substrings"
  );
});
