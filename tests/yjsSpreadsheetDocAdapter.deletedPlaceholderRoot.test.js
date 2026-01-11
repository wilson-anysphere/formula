import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("Yjs doc adapter: encode/restore tolerate empty placeholder roots with only deleted items", () => {
  // Create a doc where the comments root is an Array, then insert+delete an item
  // so the final visible state is empty but the CRDT retains tombstones.
  const source = new Y.Doc();
  const comments = source.getArray("comments");
  const comment = new Y.Map();
  comment.set("id", "c1");
  comments.push([comment]);
  comments.delete(0, 1);

  const update = Y.encodeStateAsUpdate(source);

  // Apply the update into a fresh doc *without* instantiating the comments root.
  // This commonly occurs when providers (e.g. y-websocket) apply remote updates
  // from a different module instance.
  const doc = new Y.Doc();
  Y.applyUpdate(doc, update);

  const existing = doc.share.get("comments");
  assert.ok(existing, "expected comments root placeholder to exist after applying update");

  // Create an excluded root so encodeState() takes the filtered snapshot path.
  doc.getMap("versions").set("v-local", 1);

  const adapter = createYjsSpreadsheetDocAdapter(doc, { excludeRoots: ["versions", "versionsMeta"] });

  assert.doesNotThrow(() => {
    adapter.encodeState();
  });

  // Restore from a blank snapshot should not throw even though the current doc
  // still has the placeholder root.
  const blank = Y.encodeStateAsUpdate(new Y.Doc());
  assert.doesNotThrow(() => {
    adapter.applyState(blank);
  });
});
