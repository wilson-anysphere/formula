import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { getArrayRoot } from "@formula/collab-yjs-utils";

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

