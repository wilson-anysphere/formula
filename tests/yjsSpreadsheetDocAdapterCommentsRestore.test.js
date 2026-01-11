import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("Yjs doc adapter restore rehydrates comments map even if target doc hasn't instantiated it", () => {
  const source = new Y.Doc();
  const comments = source.getMap("comments");
  const comment = new Y.Map();
  comment.set("id", "c1");
  comment.set("cellRef", "A1");
  comment.set("content", "Original comment");
  comments.set("c1", comment);
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  assert.equal(target.share.has("comments"), false);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  const restored = target.getMap("comments");
  assert.equal(restored.size, 1);
  const restoredComment = restored.get("c1");
  assert.ok(restoredComment instanceof Y.Map);
  assert.equal(restoredComment.get("content"), "Original comment");
});

test("Yjs doc adapter restore rehydrates comments array even if target doc hasn't instantiated it", () => {
  const source = new Y.Doc();
  const comments = source.getArray("comments");
  const comment = new Y.Map();
  comment.set("id", "c1");
  comment.set("cellRef", "A1");
  comment.set("content", "Original comment");
  comments.push([comment]);
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  assert.equal(target.share.has("comments"), false);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  const restored = target.getArray("comments");
  assert.equal(restored.length, 1);
  const restoredComment = restored.get(0);
  assert.ok(restoredComment instanceof Y.Map);
  assert.equal(restoredComment.get("content"), "Original comment");
});

