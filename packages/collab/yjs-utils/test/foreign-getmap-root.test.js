import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getMapRoot } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

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
