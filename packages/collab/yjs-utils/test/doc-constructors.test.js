import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { getDocTypeConstructors } from "@formula/collab-yjs-utils";

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

