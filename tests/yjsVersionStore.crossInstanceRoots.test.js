import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

test("YjsVersionStore: can read and append versions when the roots were created by a different Yjs module instance (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const versions = remote.getMap("versions");
  remote.transact(() => {
    const record = new Ycjs.Map();
    record.set("schemaVersion", 1);
    record.set("id", "v1");
    record.set("kind", "snapshot");
    record.set("timestampMs", 1);
    record.set("userId", null);
    record.set("userName", null);
    record.set("description", "from-cjs");
    record.set("checkpointName", null);
    record.set("checkpointLocked", null);
    record.set("checkpointAnnotations", null);
    record.set("compression", "none");
    record.set("snapshotEncoding", "base64");
    record.set("snapshotBase64", Buffer.from([1, 2, 3]).toString("base64"));
    versions.set("v1", record);

    const meta = remote.getMap("versionsMeta");
    const order = new Ycjs.Array();
    order.push(["v1"]);
    meta.set("order", order);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply update using CJS Yjs to simulate y-websocket behavior.
  Ycjs.applyUpdate(doc, update);

  const store = new YjsVersionStore({ doc, chunkSize: 2, compression: "none" });

  const before = await store.listVersions();
  assert.equal(before.length, 1);
  assert.equal(before[0]?.id, "v1");
  assert.equal(before[0]?.description, "from-cjs");
  assert.deepEqual(Array.from(before[0]?.snapshot ?? []), [1, 2, 3]);

  await store.saveVersion({
    id: "v2",
    kind: "snapshot",
    timestampMs: 2,
    userId: null,
    userName: null,
    description: "added-locally",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([9, 8, 7, 6]),
  });

  const after = await store.listVersions();
  assert.deepEqual(
    after.map((v) => v.id),
    ["v2", "v1"],
    "expected newly-added version to be visible and ordered by timestamp desc",
  );

  const loaded = await store.getVersion("v2");
  assert.ok(loaded);
  assert.equal(loaded.description, "added-locally");
  assert.deepEqual(Array.from(loaded.snapshot), [9, 8, 7, 6]);
});
