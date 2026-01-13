import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { ensureWorkbookSchema as ensureFromTs } from "../src/index.ts";

test("collab-workbook is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/collab-workbook");

  assert.equal(typeof mod.ensureWorkbookSchema, "function");
  assert.equal(typeof mod.getWorkbookRoots, "function");
  assert.equal(typeof mod.createSheetManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createMetadataManagerForSessionWithPermissions, "function");
  assert.equal(typeof mod.createNamedRangeManagerForSessionWithPermissions, "function");
  assert.equal(typeof ensureFromTs, "function");

  const doc = new Y.Doc();
  const roots = mod.ensureWorkbookSchema(doc, { defaultSheetId: "Sheet1" });
  assert.ok(roots);
  assert.equal(typeof roots.cells?.get, "function");
});
