import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("Yjs doc adapter: restoring a snapshot with an empty comments root does not throw and clears current comments", () => {
  const source = new Y.Doc();
  // Eagerly create the comments root but keep it empty.
  source.getMap("comments");
  const snapshot = Y.encodeStateAsUpdate(source);

  const target = new Y.Doc();
  const comments = target.getMap("comments");
  target.transact(() => {
    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("content", "Hello");
    comments.set("c1", comment);
  });
  assert.equal(comments.size, 1);

  const adapter = createYjsSpreadsheetDocAdapter(target);
  adapter.applyState(snapshot);

  assert.equal(target.getMap("comments").size, 0, "expected comments to be cleared by restore");
});

