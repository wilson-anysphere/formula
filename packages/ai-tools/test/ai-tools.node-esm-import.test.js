import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import {
  InMemoryWorkbook,
  TOOL_REGISTRY,
  ToolExecutor,
  formatA1Cell,
  formatA1Range,
  parseA1Cell,
  parseA1Range,
  validateToolCall,
} from "../src/index.ts";

test("ai-tools TS sources are importable under Node ESM (strip-types)", () => {
  assert.equal(typeof ToolExecutor, "function");
  assert.equal(typeof InMemoryWorkbook, "function");

  assert.equal(typeof parseA1Cell, "function");
  assert.equal(typeof formatA1Cell, "function");
  assert.equal(typeof parseA1Range, "function");
  assert.equal(typeof formatA1Range, "function");
  // AI tools normalize ranges/cells to include the default sheet prefix.
  assert.equal(formatA1Cell(parseA1Cell("B2")), "Sheet1!B2");
  assert.equal(formatA1Range(parseA1Range("A1:B2")), "Sheet1!A1:B2");

  assert.ok(TOOL_REGISTRY);
  assert.equal(typeof validateToolCall, "function");
});
