import assert from "node:assert/strict";
import test from "node:test";

test("collab-session exports permission-aware workbook manager helpers", async () => {
  const mod = await import("@formula/collab-session");

  assert.equal(typeof mod.createSheetManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createMetadataManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createNamedRangeManagerForSessionWithPermissions, "function");
});

