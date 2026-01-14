import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getArrayRoot, getMapRoot, getTextRoot } from "@formula/collab-yjs-utils";
import { requireYjsCjs } from "./require-yjs-cjs.js";

test("collab-yjs-utils: getMapRoot normalizes foreign roots created via CJS applyUpdate into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remoteCells.set("foo", "bar");
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply updates using the CJS build (simulating y-websocket applying updates).
  Ycjs.applyUpdate(doc, update);

  const existing = doc.share.get("cells");
  assert.ok(existing, "expected cells root to exist after applyUpdate");
  assert.equal(existing instanceof Y.Map, false, "expected cells root to be created by a foreign Yjs module instance");
  // Depending on Yjs internals, applying updates from a foreign module instance
  // can either create a foreign type (which causes `doc.getMap` to throw) or a
  // local AbstractType placeholder (which `doc.getMap` can transparently
  // convert). Both cases should be handled by `getMapRoot`.

  const cells = getMapRoot(doc, "cells");
  assert.ok(cells instanceof Y.Map, "expected getMapRoot to normalize to local Y.Map constructor");
  assert.equal(cells.get("foo"), "bar");
  assert.ok(doc.getMap("cells") instanceof Y.Map);
});

test("collab-yjs-utils: getArrayRoot normalizes foreign roots created via CJS applyUpdate into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteItems = remote.getArray("items");
  remoteItems.push([1, 2]);
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply updates using the CJS build (simulating y-websocket applying updates).
  Ycjs.applyUpdate(doc, update);

  const items = getArrayRoot(doc, "items");
  assert.ok(items instanceof Y.Array, "expected getArrayRoot to normalize to local Y.Array constructor");
  assert.deepEqual(items.toArray(), [1, 2]);
  assert.ok(doc.getArray("items") instanceof Y.Array);
});

test("collab-yjs-utils: getTextRoot normalizes foreign roots created via CJS applyUpdate into an ESM Doc", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteTitle = remote.getText("title");
  remoteTitle.insert(0, "hello");
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply updates using the CJS build (simulating y-websocket applying updates).
  Ycjs.applyUpdate(doc, update);

  const title = getTextRoot(doc, "title");
  assert.ok(title instanceof Y.Text, "expected getTextRoot to normalize to local Y.Text constructor");
  assert.equal(title.toString(), "hello");
  assert.ok(doc.getText("title") instanceof Y.Text);
  assert.equal(doc.getText("title").toString(), "hello");
});
