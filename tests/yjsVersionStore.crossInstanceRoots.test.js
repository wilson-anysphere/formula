import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";

test("YjsVersionStore: can read versions when the versions root was created by a different Yjs module instance (CJS applyUpdate)", async () => {
  const require = createRequire(import.meta.url);
  // eslint-disable-next-line import/no-named-as-default-member
  const Ycjs = require("yjs");

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

  const store = new YjsVersionStore({ doc });

  const listed = await store.listVersions();
  assert.equal(listed.length, 1);
  assert.equal(listed[0]?.id, "v1");
  assert.equal(listed[0]?.description, "from-cjs");
  assert.deepEqual(Array.from(listed[0]?.snapshot ?? []), [1, 2, 3]);
});

