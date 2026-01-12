import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { createCollabUndoService } from "../../undo/src/yjs-undo-service.js";

import { getWorkbookRoots } from "../src/index.ts";

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

test("getWorkbookRoots normalizes foreign map roots even if constructor names are renamed", () => {
  const Ycjs = requireYjsCjs();
  const doc = new Y.Doc();

  // Simulate a mixed module-loader environment where another Yjs instance eagerly
  // instantiates the `cells` root before our module touches it.
  const foreignCells = Ycjs.Doc.prototype.getMap.call(doc, "cells");
  foreignCells.set("foo", "bar");

  // Simulate a bundler-renamed constructor on the foreign instance without mutating
  // global module state.
  class RenamedMap extends foreignCells.constructor {}
  Object.setPrototypeOf(foreignCells, RenamedMap.prototype);

  const roots = getWorkbookRoots(doc);
  assert.ok(roots.cells instanceof Y.Map, "expected getWorkbookRoots to normalize to local Y.Map constructor");
  assert.equal(roots.cells.get("foo"), "bar");
});

test("getWorkbookRoots detects foreign array-backed roots even if constructor names are renamed", () => {
  const Ycjs = requireYjsCjs();
  const doc = new Y.Doc();

  const foreignCells = Ycjs.Doc.prototype.getArray.call(doc, "cells");
  assert.ok(foreignCells);
  class RenamedArray extends foreignCells.constructor {}
  Object.setPrototypeOf(foreignCells, RenamedArray.prototype);

  assert.throws(
    () => getWorkbookRoots(doc),
    /expected a Y\.Map but found a Y\.Array/,
    "expected schema mismatch error (array root should not be coerced into a map)"
  );
});

test("getWorkbookRoots normalizes a foreign AbstractType placeholder root even when it passes `instanceof Y.AbstractType` checks", () => {
  const Ycjs = requireYjsCjs();
  const doc = new Y.Doc();
  // Simulate another Yjs module instance calling `Doc.get(name)` (defaulting to
  // AbstractType) on this doc. This creates a foreign root placeholder that would
  // cause `doc.getMap("cells")` (from the ESM build) to throw.
  Ycjs.Doc.prototype.get.call(doc, "cells");

  const placeholder = doc.share.get("cells");
  assert.ok(placeholder, "expected cells root placeholder to exist");
  assert.notEqual(
    placeholder.constructor,
    Y.AbstractType,
    "expected cells root placeholder to be created by a foreign Yjs module instance"
  );

  // Sanity check: local getMap would throw "different constructor" when the placeholder
  // was created by a different Yjs module instance.
  assert.throws(() => doc.getMap("cells"), /different constructor/);

  // Patch foreign prototype chains (mirrors collab undo's behavior) so the foreign
  // placeholder passes `instanceof Y.AbstractType` checks.
  //
  // Without the `constructor === Y.AbstractType` guard in `getMapRoot`, this would
  // cause `getWorkbookRoots` to call `doc.getMap("cells")` and throw.
  createCollabUndoService({ doc, scope: placeholder }).undoManager.destroy();
  assert.equal(placeholder instanceof Y.AbstractType, true);

  const roots = getWorkbookRoots(doc);
  assert.ok(roots.cells instanceof Y.Map);
  assert.ok(doc.getMap("cells") instanceof Y.Map);

  doc.destroy();
});
