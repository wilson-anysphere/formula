import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getMapRoot } from "@formula/collab-yjs-utils";
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
