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

test("collab-yjs-utils: getMapRoot normalizes foreign AbstractType placeholder roots created via CJS Doc.get into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance touching the root via Doc.get(name),
  // leaving a foreign AbstractType placeholder under the same key.
  Ycjs.Doc.prototype.get.call(doc, "cells");

  assert.ok(doc.share.get("cells"));
  assert.throws(() => doc.getMap("cells"), /different constructor/);

  const cells = getMapRoot(doc, "cells");
  assert.ok(cells instanceof Y.Map);
  assert.ok(doc.getMap("cells") instanceof Y.Map);
});

test("collab-yjs-utils: getMapRoot normalizes foreign placeholders even when they pass `instanceof Y.AbstractType` checks", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Foreign placeholder.
  Ycjs.Doc.prototype.get.call(doc, "cells");
  const placeholder = doc.share.get("cells");
  assert.ok(placeholder);

  // Simulate collab undo's prototype patching behavior: foreign placeholders may
  // pass `instanceof Y.AbstractType` checks while still failing constructor
  // identity checks.
  const ctor = placeholder.constructor;
  assert.equal(typeof ctor, "function");
  class RenamedForeignAbstractType extends ctor {}
  Object.setPrototypeOf(RenamedForeignAbstractType.prototype, Y.AbstractType.prototype);
  Object.setPrototypeOf(placeholder, RenamedForeignAbstractType.prototype);
  assert.equal(placeholder instanceof Y.AbstractType, true);

  assert.throws(() => doc.getMap("cells"), /different constructor/);

  const cells = getMapRoot(doc, "cells");
  assert.ok(cells instanceof Y.Map);
  assert.ok(doc.getMap("cells") instanceof Y.Map);
});
