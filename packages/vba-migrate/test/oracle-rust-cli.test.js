import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import test from "node:test";

import { Workbook } from "../src/workbook.js";
import { RustCliOracle } from "../src/vba/oracle.js";

function spawnWithInput(command, args, { cwd, input }) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: ["pipe", "pipe", "pipe"] });
    const stdout = [];
    const stderr = [];
    child.stdout.on("data", (chunk) => stdout.push(chunk));
    child.stderr.on("data", (chunk) => stderr.push(chunk));
    child.on("error", reject);
    child.on("close", (code) => {
      resolve({
        code: code ?? 0,
        stdout: Buffer.concat(stdout).toString("utf8"),
        stderr: Buffer.concat(stderr).toString("utf8"),
      });
    });
    child.stdin.end(input);
  });
}

test("Rust oracle CLI emits deterministic JSON for a simple macro", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1", { makeActive: true });

  const module = {
    name: "Module1",
    code: `
Sub Main()
  Range("A1").Value = 1
  Cells(1, 2).Value = 2
  Range("A3").Formula = "=A1+B1"
End Sub
`.trim(),
  };

  const bytes = workbook.toBytes({ vbaModules: [module] });
  const oracle = new RustCliOracle();

  // Ensure the oracle is built before invoking the binary directly.
  await oracle.runMacro({ workbookBytes: bytes, macroName: "Main", inputs: [] });

  // Run twice and ensure stdout is byte-for-byte identical.
  const runOnce = async () => {
    const result = await spawnWithInput(oracle.binPath, ["run", "--macro", "Main"], {
      cwd: oracle.repoRoot,
      input: bytes,
    });
    if (result.code !== 0) {
      throw new Error(`oracle CLI exited ${result.code}: ${result.stderr}`);
    }
    return result.stdout.trim();
  };

  const out1 = await runOnce();
  const out2 = await runOnce();
  assert.equal(out1, out2);

  const report = JSON.parse(out1);
  assert.equal(report.ok, true);
  assert.equal(report.macroName, "Main");
  assert.equal(report.cellDiffs.Sheet1.A1.after.value, 1);
  assert.equal(report.cellDiffs.Sheet1.B1.after.value, 2);
  assert.equal(report.cellDiffs.Sheet1.A3.after.formula, "=A1+B1");
});
