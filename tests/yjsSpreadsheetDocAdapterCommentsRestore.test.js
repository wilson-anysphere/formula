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

test("Yjs doc adapter restore preserves legacy list comments when comments root is clobbered as a map", () => {
  const legacy = new Y.Doc();
  const comments = legacy.getArray("comments");
  const legacyComment = new Y.Map();
  legacyComment.set("id", "c1");
  legacyComment.set("cellRef", "A1");
  legacyComment.set("content", "Legacy comment");
  comments.push([legacyComment]);
  const legacySnapshot = Y.encodeStateAsUpdate(legacy);

  const mixed = new Y.Doc();
  Y.applyUpdate(mixed, legacySnapshot);

  // Simulate the old bug: choose the wrong constructor first, then add a map
  // entry. Now the root contains both map entries and legacy list items.
  const mixedComments = mixed.getMap("comments");
  const newComment = new Y.Map();
  newComment.set("id", "c2");
  newComment.set("cellRef", "A2");
  newComment.set("content", "New comment");
  mixedComments.set("c2", newComment);

  const mixedSnapshot = Y.encodeStateAsUpdate(mixed);

  const target = new Y.Doc();
  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(mixedSnapshot);

  const restored = target.getMap("comments");
  assert.equal(restored.size, 2);
  assert.ok(restored.get("c1") instanceof Y.Map);
  assert.ok(restored.get("c2") instanceof Y.Map);
  assert.equal(restored.get("c1").get("content"), "Legacy comment");
  assert.equal(restored.get("c2").get("content"), "New comment");
});

test("Yjs doc adapter restore clears legacy list comments from target when restoring a canonical comments map", () => {
  const legacy = new Y.Doc();
  const legacyComments = legacy.getArray("comments");
  const legacyComment = new Y.Map();
  legacyComment.set("id", "c1");
  legacyComment.set("cellRef", "A1");
  legacyComment.set("content", "Legacy comment");
  legacyComments.push([legacyComment]);
  const legacySnapshot = Y.encodeStateAsUpdate(legacy);

  // Target doc starts in a clobbered state: comments is a map root, but contains
  // a legacy list entry (parentSub === null).
  const target = new Y.Doc();
  Y.applyUpdate(target, legacySnapshot);
  target.getMap("comments");

  const canonical = new Y.Doc();
  const canonicalComments = canonical.getMap("comments");
  const canonicalComment = new Y.Map();
  canonicalComment.set("id", "c2");
  canonicalComment.set("cellRef", "A2");
  canonicalComment.set("content", "Canonical comment");
  canonicalComments.set("c2", canonicalComment);
  const canonicalSnapshot = Y.encodeStateAsUpdate(canonical);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(canonicalSnapshot);

  const restored = target.getMap("comments");
  assert.equal(restored.size, 1);
  assert.equal(restored.has("c1"), false);
  assert.equal(restored.has("c2"), true);

  // Ensure the legacy list item is gone (no non-deleted items with parentSub null).
  let item = restored._start;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      assert.fail("expected restored comments map to have no legacy list items");
    }
    item = item.right;
  }
});

test("Yjs doc adapter restore rehydrates additional map roots present only in the snapshot", () => {
  const source = new Y.Doc();
  const settings = source.getMap("settings");
  settings.set("theme", "dark");
  settings.set("gridlines", true);
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  assert.equal(target.share.has("settings"), false);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  const restored = target.getMap("settings");
  assert.equal(restored.get("theme"), "dark");
  assert.equal(restored.get("gridlines"), true);
});

test("Yjs doc adapter restore preserves Y.Text formatting when restoring text roots", () => {
  const source = new Y.Doc();
  const note = source.getText("note");
  note.insert(0, "hello");
  note.format(0, 5, { bold: true });
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  assert.equal(target.share.has("note"), false);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  const restored = target.getText("note");
  assert.equal(restored.toString(), "hello");
  assert.deepEqual(restored.toDelta(), [{ insert: "hello", attributes: { bold: true } }]);
});

test("Yjs doc adapter restore preserves Y.Text embeds when restoring text roots", () => {
  const source = new Y.Doc();
  const note = source.getText("note");
  note.insertEmbed(0, { type: "emoji", value: "🙂" });
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  const restored = target.getText("note");
  assert.deepEqual(restored.toDelta(), [{ insert: { type: "emoji", value: "🙂" } }]);
});
