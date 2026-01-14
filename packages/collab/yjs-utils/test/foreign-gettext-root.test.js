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

test("collab-yjs-utils: getTextRoot normalizes foreign AbstractType placeholder roots created via CJS Doc.get into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance touching the root via Doc.get(name),
  // leaving a foreign AbstractType placeholder under the same key.
  Ycjs.Doc.prototype.get.call(doc, "title");

  assert.ok(doc.share.get("title"));
  assert.throws(() => doc.getText("title"), /different constructor/);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text);
  assert.ok(doc.getText("title") instanceof Y.Text);
});

test("collab-yjs-utils: getTextRoot preserves foreign formatting/content when normalizing a text root", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  const foreign = Ycjs.Doc.prototype.getText.call(doc, "title");
  foreign.insert(0, "hello");
  foreign.format(0, 5, { bold: true });

  assert.throws(() => doc.getText("title"), /different constructor/);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text);
  assert.equal(title.toString(), "hello");

  const delta = title.toDelta();
  const inserted = delta.map((op) => (typeof op.insert === "string" ? op.insert : "")).join("");
  assert.equal(inserted, "hello");
  assert.equal(
    delta.some((op) => typeof op.attributes === "object" && op.attributes && op.attributes.bold === true),
    true,
  );
});

test("collab-yjs-utils: getTextRoot preserves foreign embeds when normalizing a text root", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  const foreign = Ycjs.Doc.prototype.getText.call(doc, "title");
  foreign.insert(0, "hi");
  foreign.insertEmbed(2, { foo: "bar" });

  assert.throws(() => doc.getText("title"), /different constructor/);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text);

  const delta = title.toDelta();
  assert.equal(
    delta.some((op) => op && typeof op.insert === "object" && op.insert && op.insert.foo === "bar"),
    true,
  );
});

test("collab-yjs-utils: getTextRoot preserves foreign embedded Yjs types when normalizing a text root", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  const foreign = Ycjs.Doc.prototype.getText.call(doc, "title");
  const embedded = new Ycjs.Map();
  embedded.set("x", 1);
  foreign.insertEmbed(0, embedded);

  assert.throws(() => doc.getText("title"), /different constructor/);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text);

  const delta = title.toDelta();
  const insertedMap = delta.find((op) => op && op.insert && typeof op.insert.get === "function")?.insert;
  assert.ok(insertedMap, "expected embedded Yjs type to round-trip through toDelta()");
  assert.equal(insertedMap.get("x"), 1);
});
