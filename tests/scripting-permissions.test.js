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
`);

  assert.ok(result.error, "expected script to fail due to denied network");
  assert.match(result.error.message, /Network access/i);
});

test("scripting (node): allowlist mode rejects non-allowlisted hosts (fetch + WebSocket)", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(workbook);

  const allowlist = { network: "allowlist", networkAllowlist: ["localhost"] };

  const fetchResult = await runtime.run(
    `
export default async function main(ctx) {
  await ctx.fetch("https://example.com");
}
`,
    { permissions: allowlist },
  );

  assert.ok(fetchResult.error, "expected fetch to be blocked by allowlist");
  assert.match(fetchResult.error.message, /example\.com/i);

  const wsResult = await runtime.run(
    `
export default async function main() {
  new WebSocket("wss://example.com");
}
`,
    { permissions: allowlist },
  );

  assert.ok(wsResult.error, "expected WebSocket to be blocked by allowlist");
  assert.match(wsResult.error.message, /example\.com/i);
});
