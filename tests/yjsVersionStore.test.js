import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";

test("YjsVersionStore: save/get/list/update/delete (gzip + chunks)", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsVersionStore({ doc: ydoc, compression: "gzip", chunkSize: 4 });

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
    snapshot: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]),
  };

  await store.saveVersion(v1);

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

