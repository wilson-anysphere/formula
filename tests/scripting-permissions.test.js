import assert from "node:assert/strict";
import test from "node:test";

import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

test("scripting (node): blocks network access by default", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);
  const result = await runtime.run(`
 export default async function main(ctx) {
   await ctx.fetch("https://example.com");
 }
`, { timeoutMs: 30_000 });

  assert.ok(result.error, "expected script to fail due to denied network");
  assert.match(result.error.message, /Network access/i);
});

test("scripting (node): allowlist mode rejects non-allowlisted hosts (fetch + WebSocket)", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);

  const allowlist = { network: { mode: "allowlist", allowlist: ["localhost"] } };

  const fetchResult = await runtime.run(
    `
 export default async function main(ctx) {
   await ctx.fetch("https://example.com");
 }
 `,
    { permissions: allowlist, timeoutMs: 30_000 },
  );

  assert.ok(fetchResult.error, "expected fetch to be blocked by allowlist");
  assert.match(fetchResult.error.message, /example\.com/i);

  const wsResult = await runtime.run(
    `
 export default async function main() {
   new WebSocket("wss://example.com");
 }
 `,
    { permissions: allowlist, timeoutMs: 30_000 },
  );

  assert.ok(wsResult.error, "expected WebSocket to be blocked by allowlist");
  assert.match(wsResult.error.message, /WebSocket/i);
});

test("scripting (node): module scripts cannot import other modules", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);
  const result = await runtime.run(`
import { readFileSync } from "node:fs";
export default async function main(ctx) {
  ctx.ui.log(readFileSync);
}
`, { timeoutMs: 30_000 });

  assert.ok(result.error, "expected imports to be rejected");
  assert.match(result.error.message, /Imports are not supported/i);
  assert.match(result.error.message, /node:fs/i);
});

test("scripting (node): script-body cannot use dynamic import()", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);
  const result = await runtime.run(`
const mod = await import("node:fs");
ctx.ui.log(mod);
`, { timeoutMs: 30_000 });

  assert.ok(result.error, "expected dynamic import to be rejected");
  assert.match(result.error.message, /dynamic import/i);
  assert.match(result.error.message, /node:fs/i);
});
