import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";
import { requireYjsCjs } from "../../../collab/yjs-utils/test/require-yjs-cjs.js";
import { patchForeignAbstractTypeConstructor } from "../../../collab/yjs-utils/src/index.ts";

import { YjsVersionStore } from "./yjsVersionStore.js";

test("YjsVersionStore normalizes foreign versioning roots created by a different Yjs instance (CJS getMap)", async () => {
  const Ycjs = requireYjsCjs();
  const doc = new Y.Doc();

  // Simulate a mixed module loader environment where another Yjs instance eagerly
  // instantiates the version history roots before YjsVersionStore is constructed.
  Ycjs.Doc.prototype.getMap.call(doc, "versions");
  Ycjs.Doc.prototype.getMap.call(doc, "versionsMeta");

  assert.throws(() => doc.getMap("versions"), /different constructor/);
  assert.throws(() => doc.getMap("versionsMeta"), /different constructor/);

  const store = new YjsVersionStore({
    doc,
    // Use defaults: snapshotEncoding="chunks". Without root normalization, the store
    // would fall back to base64 because it can't infer a compatible Y.Array ctor.
    snapshotEncoding: "chunks",
    writeMode: "single",
  });

  // Root normalization should re-wrap foreign roots into the local Yjs instance so
  // other code can safely call `doc.getMap("versions")`.
  assert.ok(doc.share.get("versions") instanceof Y.Map);
  assert.ok(doc.share.get("versionsMeta") instanceof Y.Map);
  assert.ok(doc.getMap("versions") instanceof Y.Map);
  assert.ok(doc.getMap("versionsMeta") instanceof Y.Map);

  const snapshot = new Uint8Array([1, 2, 3]);
  await store.saveVersion({
    id: "v1",
    kind: "snapshot",
    timestampMs: 1,
    userId: null,
    userName: null,
    description: null,
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot,
  });

  const raw = doc.getMap("versions").get("v1");
  assert.ok(raw instanceof Y.Map);
  assert.equal(raw.get("snapshotEncoding"), "chunks");
  assert.ok(raw.get("snapshotChunks") instanceof Y.Array);
  assert.equal(raw.get("snapshotBase64"), undefined);

  const roundTrip = await store.getVersion("v1");
  assert.ok(roundTrip);
  assert.deepEqual(Array.from(roundTrip.snapshot), Array.from(snapshot));
});

test("YjsVersionStore normalizes foreign AbstractType placeholder roots even when they pass instanceof checks", async () => {
  const Ycjs = requireYjsCjs();
  const doc = new Y.Doc();

  // Simulate another Yjs module instance touching the roots via Doc.get(name),
  // leaving a foreign AbstractType placeholder under the same key.
  Ycjs.Doc.prototype.get.call(doc, "versions");
  Ycjs.Doc.prototype.get.call(doc, "versionsMeta");

  const placeholder = doc.share.get("versions");
  assert.ok(placeholder);
  assert.throws(() => doc.getMap("versions"), /different constructor/);

  // Patch prototype chain so the foreign placeholder passes `instanceof Y.AbstractType`
  // checks (mirrors collab undo's prototype patching behavior).
  patchForeignAbstractTypeConstructor(placeholder);
  assert.equal(placeholder instanceof Y.AbstractType, true);

  const store = new YjsVersionStore({
    doc,
    snapshotEncoding: "chunks",
    writeMode: "single",
  });

  assert.ok(doc.getMap("versions") instanceof Y.Map);
  assert.ok(doc.getMap("versionsMeta") instanceof Y.Map);

  const snapshot = new Uint8Array([1, 2, 3]);
  await store.saveVersion({
    id: "v1",
    kind: "snapshot",
    timestampMs: 1,
    userId: null,
    userName: null,
    description: null,
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot,
  });

  const roundTrip = await store.getVersion("v1");
  assert.ok(roundTrip);
  assert.deepEqual(Array.from(roundTrip.snapshot), Array.from(snapshot));
});
