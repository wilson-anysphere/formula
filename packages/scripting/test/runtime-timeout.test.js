import assert from "node:assert/strict";
import test from "node:test";

import { ScriptRuntime, Workbook } from "@formula/scripting/node";

test("ScriptRuntime (node) times out scripts that never resolve", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);

  const start = Date.now();
  const result = await runtime.run(`await new Promise(() => {});`, { timeoutMs: 200 });
  const elapsed = Date.now() - start;

  assert.ok(elapsed < 2_000, `expected ScriptRuntime.run to time out quickly, took ${elapsed}ms`);
  assert.ok(result.error, "expected a timeout error");
  assert.match(result.error.message, /timed out/i);
});

