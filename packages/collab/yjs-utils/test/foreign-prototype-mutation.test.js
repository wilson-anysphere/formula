import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getTextRoot } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: normalizing foreign roots does not mutate the foreign module's global Content* constructors", () => {
  const Ycjs = requireYjsCjs();

  // Capture a foreign ContentString constructor reference.
  const foreignDoc1 = new Ycjs.Doc();
  const foreignText1 = foreignDoc1.getText("t");
  foreignText1.insert(0, "hello");
  assert.equal(foreignText1.toString(), "hello");

  const item1 = /** @type {any} */ (foreignText1)._start;
  assert.ok(item1, "expected internal Item struct");
  const foreignContentStringCtor = item1.content?.constructor;
  assert.equal(typeof foreignContentStringCtor, "function", "expected ContentString constructor to be a function");

  // Apply an update created by the foreign module instance into a local ESM doc.
  const update = Ycjs.encodeStateAsUpdate(foreignDoc1);
  const doc = new Y.Doc();
  Ycjs.applyUpdate(doc, update);

  const normalized = getTextRoot(doc, "t");
  assert.ok(normalized instanceof Y.Text);
  assert.equal(normalized.toString(), "hello");

  // Regression guard: `getTextRoot`/normalization must not patch foreign constructors/prototypes
  // globally. If it did, subsequent docs created via the foreign module instance would have
  // broken Y.Text constructor equality checks and `toString()` would return empty.
  const foreignDoc2 = new Ycjs.Doc();
  const foreignText2 = foreignDoc2.getText("t");
  foreignText2.insert(0, "world");

  assert.equal(foreignText2.toString(), "world");

  const item2 = /** @type {any} */ (foreignText2)._start;
  assert.ok(item2, "expected internal Item struct");
  assert.equal(item2.content?.constructor, foreignContentStringCtor);
  assert.equal(foreignContentStringCtor.prototype.constructor, foreignContentStringCtor);
});

