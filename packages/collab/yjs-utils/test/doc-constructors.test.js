import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getDocTypeConstructors } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: getDocTypeConstructors returns local constructors for an ESM doc", () => {
  const doc = new Y.Doc();
  const ctors = getDocTypeConstructors(doc);
  assert.equal(ctors.Map, Y.Map);
  assert.equal(ctors.Array, Y.Array);
  assert.equal(ctors.Text, Y.Text);
});

test("collab-yjs-utils: getDocTypeConstructors returns foreign constructors for a CJS doc", () => {
  const Ycjs = requireYjsCjs();
  const doc = new Ycjs.Doc();
  const beforeShareSize = doc.share.size;

  const ctors = getDocTypeConstructors(doc);

  // Should not mutate the original doc.
  assert.equal(doc.share.size, beforeShareSize);

  const map = new ctors.Map();
  const arr = new ctors.Array();
  const text = new ctors.Text();

  assert.equal(map instanceof Y.Map, false);
  assert.equal(arr instanceof Y.Array, false);
  assert.equal(text instanceof Y.Text, false);

  assert.equal(map instanceof Ycjs.Map, true);
  assert.equal(arr instanceof Ycjs.Array, true);
  assert.equal(text instanceof Ycjs.Text, true);
});
