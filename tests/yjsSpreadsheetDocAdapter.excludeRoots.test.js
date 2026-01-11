import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

test("Yjs doc adapter: excludeRoots prevents version-history updates from marking the workbook dirty", () => {
  const doc = new Y.Doc();
  const adapter = createYjsSpreadsheetDocAdapter(doc, { excludeRoots: ["versions", "versionsMeta"] });

  let updates = 0;
  adapter.on("update", () => {
    updates += 1;
  });

  const cells = doc.getMap("cells");

  // Root update: creating a cell should count as a workbook update.
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", 1);
    cells.set("Sheet1:0:0", cell);
  });
  assert.equal(updates, 1);

  // Deep update: editing the nested cell map should still count as a workbook update.
  doc.transact(() => {
    const cell = cells.get("Sheet1:0:0");
    assert.ok(cell instanceof Y.Map);
    cell.set("value", 2);
  });
  assert.equal(updates, 2);

  // Excluded root update: version history writes should NOT count as workbook updates.
  const versions = doc.getMap("versions");
  doc.transact(() => {
    const record = new Y.Map();
    record.set("kind", "checkpoint");
    versions.set("v1", record);
  });
  assert.equal(updates, 2);

  // Excluded deep update: editing the nested version record should also be ignored.
  doc.transact(() => {
    const record = versions.get("v1");
    assert.ok(record instanceof Y.Map);
    record.set("checkpointLocked", true);
  });
  assert.equal(updates, 2);
});

