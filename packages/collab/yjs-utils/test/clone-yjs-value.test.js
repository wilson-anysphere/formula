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
  assert.equal(cloned.get("x"), 1);
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
