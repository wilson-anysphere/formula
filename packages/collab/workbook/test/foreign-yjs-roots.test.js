import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

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
