import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("restores_snapshot_roots_not_instantiated_in_current_doc", () => {
  const docA = new Y.Doc();
  const commentsA = docA.getMap("comments");

  docA.transact(() => {
    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("content", "Hello");
    commentsA.set("c1", comment);
  });

  const snapshot = Y.encodeStateAsUpdate(docA);

  const docB = new Y.Doc();
  assert.equal(docB.share.has("comments"), false);

  const adapterB = createYjsSpreadsheetDocAdapter(docB);
  adapterB.applyState(snapshot);

  assert.equal(docB.share.has("comments"), true);
  const commentsB = docB.getMap("comments");
  assert.equal(commentsB.size, 1);

  const restoredComment = commentsB.get("c1");
  assert.ok(restoredComment instanceof Y.Map);
  assert.equal(restoredComment.get("content"), "Hello");
});

test("throws_on_root_kind_mismatch", () => {
  const docA = new Y.Doc();
  docA.getMap("foo").set("bar", 1);
  const snapshot = Y.encodeStateAsUpdate(docA);

  const docB = new Y.Doc();
  docB.getArray("foo").push([1, 2, 3]);

  const adapterB = createYjsSpreadsheetDocAdapter(docB);
  assert.throws(() => adapterB.applyState(snapshot), /schema mismatch.*foo/i);
});

