import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getArrayRoot } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: getArrayRoot normalizes foreign roots created via CJS getArray into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance eagerly instantiating an array root.
  const foreign = Ycjs.Doc.prototype.getArray.call(doc, "items");
  foreign.push(["a", "b"]);

  assert.throws(() => doc.getArray("items"), /different constructor/);

  const items = getArrayRoot(doc, "items");
  assert.ok(items instanceof Y.Array);
  assert.deepEqual(items.toArray(), ["a", "b"]);
  assert.ok(doc.getArray("items") instanceof Y.Array);
});

test("collab-yjs-utils: getArrayRoot normalizes foreign AbstractType placeholder roots created via CJS Doc.get into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance touching the root via Doc.get(name),
  // leaving a foreign AbstractType placeholder under the same key.
  Ycjs.Doc.prototype.get.call(doc, "items");

  assert.ok(doc.share.get("items"));
  assert.throws(() => doc.getArray("items"), /different constructor/);

  const items = getArrayRoot(doc, "items");
  assert.ok(items instanceof Y.Array);
  assert.ok(doc.getArray("items") instanceof Y.Array);
});
