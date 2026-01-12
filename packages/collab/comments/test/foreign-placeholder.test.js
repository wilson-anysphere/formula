import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { CommentManager, getCommentsRoot } from "../src/manager.ts";

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

test('collab comments: tolerates foreign AbstractType placeholder root created via CJS Doc.get("comments")', () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate a mixed-module loader environment where another Yjs module instance
  // uses `Doc.get(name)` (defaulting to `AbstractType`) to touch the `comments` root.
  Ycjs.Doc.prototype.get.call(doc, "comments");

  // The root exists, but the placeholder is not an instanceof the local Y.AbstractType.
  assert.ok(doc.share.get("comments"));
  assert.equal(doc.share.get("comments") instanceof Y.AbstractType, false);

  // Regression: `getCommentsRoot` should not throw "different constructor".
  const root = getCommentsRoot(doc);
  assert.equal(root.kind, "map");
  assert.ok(root.kind === "map" && root.map instanceof Y.Map);

  // Ensure the doc can now safely instantiate the canonical map root.
  assert.ok(doc.getMap("comments") instanceof Y.Map);

  // CommentManager should function normally.
  const mgr = new CommentManager(doc);
  mgr.addComment({
    id: "c1",
    cellRef: "A1",
    kind: "threaded",
    content: "Hello",
    author: { id: "u1", name: "Alice" },
    now: 1,
  });
  assert.deepEqual(mgr.listAll().map((c) => c.id), ["c1"]);

  doc.destroy();
});

