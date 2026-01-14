import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { cloneYjsValue } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: cloneYjsValue preserves source constructors when constructors are omitted", () => {
  const Ycjs = requireYjsCjs();

  const foreignMap = new Ycjs.Map();
  foreignMap.set("x", 1);

  const cloned = cloneYjsValue(foreignMap);
  assert.equal(cloned instanceof Y.Map, false);
  assert.equal(cloned instanceof Ycjs.Map, true);

  // Newer Yjs versions only allow reading content once the type is integrated
  // into a document. Insert the cloned map into a doc created with the same
  // module instance so we can verify contents without constructor mismatches.
  const doc = new Ycjs.Doc();
  const root = doc.getMap("root");
  root.set("data", cloned);

  const data = root.get("data");
  assert.equal(data instanceof Ycjs.Map, true);
  assert.equal(data.get("x"), 1);
});

test("collab-yjs-utils: cloneYjsValue can clone foreign values into local constructors for insertion", () => {
  const Ycjs = requireYjsCjs();

  const foreign = new Ycjs.Map();
  foreign.set("k", "v");

  const foreignText = new Ycjs.Text();
  foreignText.insert(0, "hello");
  foreign.set("t", foreignText);

  const foreignArr = new Ycjs.Array();
  foreignArr.push([1, 2]);
  foreign.set("a", foreignArr);

  const cloned = cloneYjsValue(foreign, { Map: Y.Map, Array: Y.Array, Text: Y.Text });
  assert.ok(cloned instanceof Y.Map);

  const doc = new Y.Doc();
  const root = doc.getMap("root");
  root.set("data", cloned);

  const data = root.get("data");
  assert.ok(data instanceof Y.Map);
  assert.equal(data.get("k"), "v");

  const t = data.get("t");
  assert.ok(t instanceof Y.Text);
  assert.equal(t.toString(), "hello");

  const a = data.get("a");
  assert.ok(a instanceof Y.Array);
  assert.deepEqual(a.toArray(), [1, 2]);
});

test("collab-yjs-utils: cloneYjsValue can clone unintegrated foreign Y.Text with formatting and embeds", () => {
  const Ycjs = requireYjsCjs();

  const foreignText = new Ycjs.Text();
  foreignText.insert(0, "hello");
  foreignText.format(0, 5, { bold: true });
  foreignText.insertEmbed(5, { foo: "bar" });
  const embedded = new Ycjs.Map();
  embedded.set("x", 1);
  foreignText.insertEmbed(6, embedded);

  const cloned = cloneYjsValue(foreignText, { Map: Y.Map, Array: Y.Array, Text: Y.Text });
  assert.ok(cloned instanceof Y.Text);

  const doc = new Y.Doc();
  const root = doc.getMap("root");
  root.set("t", cloned);

  const t = root.get("t");
  assert.ok(t instanceof Y.Text);

  const delta = t.toDelta();
  const inserted = delta.map((op) => (typeof op.insert === "string" ? op.insert : "")).join("");
  assert.equal(inserted, "hello");
  assert.equal(
    delta.some((op) => typeof op.attributes === "object" && op.attributes && op.attributes.bold === true),
    true,
  );
  assert.equal(
    delta.some((op) => op && typeof op.insert === "object" && op.insert && op.insert.foo === "bar"),
    true,
  );

  const insertedMap = delta.find((op) => op && op.insert && typeof op.insert.get === "function")?.insert;
  assert.ok(insertedMap, "expected embedded Yjs type to round-trip through toDelta()");
  assert.equal(insertedMap.get("x"), 1);
});
