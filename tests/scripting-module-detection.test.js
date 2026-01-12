import assert from "node:assert/strict";
import test from "node:test";

import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

test("scripting: script-body is not treated as module when export appears only in a comment", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);
  const result = await runtime.run(`// export default async function main(ctx) {}
 ctx.ui.log("ok");
`, { timeoutMs: 30_000 });

  assert.equal(result.error, undefined, result.error?.message);
  assert.ok(result.logs.some((entry) => entry.message.includes("ok")), "expected console logs to include 'ok'");
});

test("scripting: module script without default export fails with clear error", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);
  const result = await runtime.run(`export const value = 123;`, { timeoutMs: 30_000 });

  assert.ok(result.error, "expected module script to fail without default export");
  assert.match(result.error.message, /export.*default|default function|must export/i);
});
