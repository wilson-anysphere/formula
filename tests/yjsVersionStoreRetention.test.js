import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";
import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";
import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";

test("VersionManager retention works with YjsVersionStore (history in-doc, excludeRoots)", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");

  const store = new YjsVersionStore({ doc: ydoc });
  const doc = createYjsSpreadsheetDocAdapter(ydoc, { excludeRoots: ["versions", "versionsMeta"] });
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    retention: { maxSnapshots: 2 },
  });

  for (let i = 0; i < 5; i += 1) {
    ydoc.transact(() => {
      const cell = new Y.Map();
      cell.set("value", i);
      cells.set(`Sheet1:0:${i}`, cell);
    });
    await vm.createSnapshot({ description: `s${i}` });
  }

  const versions = await vm.listVersions();
  const snapshots = versions.filter((v) => v.kind === "snapshot");
  assert.equal(snapshots.length, 2, "expected retention to prune older snapshots");
  assert.deepEqual(
    snapshots.map((v) => v.description),
    ["s4", "s3"],
    "expected newest snapshots to be retained"
  );
});

