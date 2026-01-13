import assert from "node:assert/strict";
import test from "node:test";

import * as browser from "./browser.js";

test("branches browser entrypoint exports the browser-safe surface (no SQLiteBranchStore)", () => {
  assert.equal(typeof browser.BranchService, "function");
  assert.equal(typeof browser.YjsBranchStore, "function");
  assert.equal(typeof browser.yjsDocToDocumentState, "function");
  assert.equal(typeof browser.applyDocumentStateToYjsDoc, "function");
  assert.equal(typeof browser.rowColToA1, "function");
  assert.equal(typeof browser.a1ToRowCol, "function");

  // Pure helpers (safe).
  assert.equal(typeof browser.mergeDocumentStates, "function");
  assert.equal(typeof browser.diffDocumentStates, "function");
  assert.equal(typeof browser.normalizeDocumentState, "function");

  // Node-only store must never be exported from the browser entrypoint.
  assert.equal("SQLiteBranchStore" in browser, false);
});
