import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
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

test("ai-tools TS sources are importable under Node ESM when executing TS sources directly", () => {
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

test("ai-tools package entrypoint is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/ai-tools");

  assert.equal(typeof mod.ToolExecutor, "function");
  assert.equal(typeof mod.InMemoryWorkbook, "function");
  assert.equal(typeof mod.parseA1Cell, "function");
  assert.equal(typeof mod.formatA1Cell, "function");
  assert.equal(typeof mod.parseA1Range, "function");
  assert.equal(typeof mod.formatA1Range, "function");
  assert.equal(typeof mod.validateToolCall, "function");
  assert.ok(mod.TOOL_REGISTRY);
});
