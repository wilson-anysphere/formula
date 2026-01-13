import assert from "node:assert/strict";
import test from "node:test";

import * as mod from "@formula/collab-session";

test("collab-session exports permission-aware workbook manager helpers", () => {
  assert.equal(typeof mod.createSheetManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createMetadataManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createNamedRangeManagerForSessionWithPermissions, "function");
});
