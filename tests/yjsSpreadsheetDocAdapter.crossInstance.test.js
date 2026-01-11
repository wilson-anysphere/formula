import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("Yjs doc adapter: works when the Y.Doc comes from a different Yjs module instance (CJS vs ESM)", () => {
  const require = createRequire(import.meta.url);
  // y-websocket pulls in the CJS build of Yjs; in pnpm workspaces it's possible
  // for the app to end up with both ESM + CJS module instances. This test
  // ensures our adapter doesn't rely on `instanceof` checks.
  // eslint-disable-next-line import/no-named-as-default-member
  const Ycjs = require("yjs");

  const source = new Ycjs.Doc();
  const sourceCells = source.getMap("cells");
  source.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", "alpha");
    sourceCells.set("Sheet1:0:0", cell);
  });
  const snapshot = Ycjs.encodeStateAsUpdate(source);

  const target = new Ycjs.Doc();
  target.getMap("cells").set("Sheet1:0:0", new Ycjs.Map());

  // Create excluded roots so encodeState() goes through the "filtered snapshot"
  // path that clones values between documents.
  target.getMap("versions").set("v-local", new Ycjs.Map());

  const adapter = createYjsSpreadsheetDocAdapter(target, { excludeRoots: ["versions", "versionsMeta"] });
  adapter.applyState(snapshot);

  const restoredCells = target.getMap("cells");
  const restoredCell = restoredCells.get("Sheet1:0:0");
  assert.ok(restoredCell, "expected restored cell map to exist");
  assert.equal(restoredCell.get("value"), "alpha");

  // encodeState() should exclude the version-history roots even though they
  // exist in the underlying doc.
  const filteredSnapshot = adapter.encodeState();
  const replay = new Y.Doc();
  Y.applyUpdate(replay, filteredSnapshot);

  assert.equal(replay.share.has("versions"), false, "expected versions root to be excluded");
  assert.equal(replay.getMap("cells").get("Sheet1:0:0")?.get("value"), "alpha");
});

