import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { YjsVersionStore } from "./yjsVersionStore.js";

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

