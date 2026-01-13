import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";

test("YjsVersionStore: save/get/list/update/delete (gzip + chunks)", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsVersionStore({ doc: ydoc, compression: "gzip", chunkSize: 4 });

  const snapshotBytes = new Uint8Array(256);
  for (let i = 0; i < snapshotBytes.length; i += 1) snapshotBytes[i] = i;

  const v1 = {
    id: "v1",
    kind: "checkpoint",
    timestampMs: 100,
    userId: "u1",
    userName: "User",
    description: "Approved",
    checkpointName: "Approved",
    checkpointLocked: false,
    checkpointAnnotations: null,
    snapshot: snapshotBytes,
  };

  await store.saveVersion(v1);

  {
    const record = ydoc.getMap("versions").get("v1");
    const chunks = record?.get?.("snapshotChunks");
    assert.ok(chunks, "expected snapshotChunks to be stored");
    assert.ok(typeof chunks.toArray === "function");
    assert.ok(chunks.toArray().length > 1, "expected snapshot to be chunked");
  }

  const loaded1 = await store.getVersion("v1");
  assert.ok(loaded1);
  assert.equal(loaded1.kind, "checkpoint");
  assert.equal(loaded1.checkpointLocked, false);
  assert.deepEqual(Array.from(loaded1.snapshot), Array.from(v1.snapshot));

  await store.updateVersion("v1", { checkpointLocked: true });
  const updated1 = await store.getVersion("v1");
  assert.ok(updated1);
  assert.equal(updated1.checkpointLocked, true);

  const v2 = {
    ...v1,
    id: "v2",
    kind: "snapshot",
    timestampMs: 200,
    description: "Auto-save",
  };
  await store.saveVersion(v2);

  const listed = await store.listVersions();
  assert.deepEqual(
    listed.map((v) => v.id),
    ["v2", "v1"],
    "expected listVersions to sort by timestamp desc"
  );

  await store.deleteVersion("v2");
  const afterDelete = await store.listVersions();
  assert.deepEqual(afterDelete.map((v) => v.id), ["v1"]);
});

test("YjsVersionStore: base64 encoding round-trips", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsVersionStore({ doc: ydoc, snapshotEncoding: "base64", compression: "none" });

  const v1 = {
    id: "v1",
    kind: "snapshot",
    timestampMs: 1,
    userId: null,
    userName: null,
    description: "hello",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([0, 255, 1, 2, 3]),
  };

  await store.saveVersion(v1);

  const record = ydoc.getMap("versions").get("v1");
  assert.equal(typeof record?.get?.("snapshotBase64"), "string", "expected snapshotBase64 to be stored");

  const loaded = await store.getVersion("v1");
  assert.ok(loaded);
  assert.deepEqual(Array.from(loaded.snapshot), Array.from(v1.snapshot));
});

test("YjsVersionStore: rejects unknown schemaVersion", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsVersionStore({ doc: ydoc });

  const versions = ydoc.getMap("versions");
  ydoc.transact(() => {
    const record = new Y.Map();
    record.set("schemaVersion", 999);
    record.set("kind", "snapshot");
    record.set("timestampMs", 1);
    record.set("compression", "none");
    record.set("snapshotEncoding", "base64");
    record.set("snapshotBase64", Buffer.from([1, 2, 3]).toString("base64"));
    versions.set("bad", record);
  });

  await assert.rejects(() => store.getVersion("bad"), /unsupported schemaVersion/i);
});

test("YjsVersionStore: listVersions is deterministic when timestamps are equal (insertion order tie-breaker)", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsVersionStore({ doc: ydoc, compression: "none" });

  await store.saveVersion({
    id: "v1",
    kind: "snapshot",
    timestampMs: 0,
    userId: null,
    userName: null,
    description: "first",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([1]),
  });

  await store.saveVersion({
    id: "v2",
    kind: "snapshot",
    timestampMs: 0,
    userId: null,
    userName: null,
    description: "second",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([2]),
  });

  const versions = await store.listVersions();
  assert.deepEqual(
    versions.map((v) => v.id),
    ["v2", "v1"],
    "expected last inserted to be listed first when timestamps match",
  );
});

test("YjsVersionStore: rejects invalid chunkSize values", async () => {
  const doc = new Y.Doc();
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, chunkSize: 0 }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, chunkSize: -1 }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, chunkSize: Number.NaN }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, chunkSize: 1.5 }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, chunkSize: Number.POSITIVE_INFINITY }), /chunkSize/i);
});

test("YjsVersionStore: rejects invalid maxChunksPerTransaction values", async () => {
  const doc = new Y.Doc();
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, maxChunksPerTransaction: 0 }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, maxChunksPerTransaction: -1 }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, maxChunksPerTransaction: Number.NaN }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsVersionStore({ doc, maxChunksPerTransaction: 1.5 }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(
    () => new YjsVersionStore({ doc, maxChunksPerTransaction: Number.POSITIVE_INFINITY }),
    /maxChunksPerTransaction/i,
  );
});
