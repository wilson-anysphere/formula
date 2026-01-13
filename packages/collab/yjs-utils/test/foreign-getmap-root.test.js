import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { getMapRoot } from "@formula/collab-yjs-utils";

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

test("collab-yjs-utils: getMapRoot normalizes foreign roots created via CJS getMap into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance eagerly instantiating a root.
  const foreign = Ycjs.Doc.prototype.getMap.call(doc, "cells");
  foreign.set("foo", "bar");

  assert.throws(() => doc.getMap("cells"), /different constructor/);

  const cells = getMapRoot(doc, "cells");
  assert.ok(cells instanceof Y.Map);
  assert.equal(cells.get("foo"), "bar");
  assert.ok(doc.getMap("cells") instanceof Y.Map);
});

