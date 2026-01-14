import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { patchForeignItemConstructor } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: patchForeignItemConstructor patches foreign Item structs to pass instanceof checks", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.getText("t").insert(0, "hello");
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply via foreign module instance to produce foreign root + foreign Item structs.
  Ycjs.applyUpdate(doc, update);

  const root = /** @type {any} */ (doc.share.get("t"));
  assert.ok(root, "expected text root");

  const item = root._start;
  assert.ok(item, "expected internal Item struct");

  // Sanity: foreign item should not be an instance of the local Y.Item constructor.
  assert.equal(item instanceof Y.Item, false);

  patchForeignItemConstructor(item);

  assert.equal(item instanceof Y.Item, true);
});
