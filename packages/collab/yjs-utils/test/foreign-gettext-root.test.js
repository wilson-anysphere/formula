import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getTextRoot } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: getTextRoot normalizes foreign roots created via CJS getText into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance eagerly instantiating a text root.
  const foreign = Ycjs.Doc.prototype.getText.call(doc, "title");
  foreign.insert(0, "hello");

  assert.throws(() => doc.getText("title"), /different constructor/);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text);
  assert.equal(title.toString(), "hello");
  assert.ok(doc.getText("title") instanceof Y.Text);
});
