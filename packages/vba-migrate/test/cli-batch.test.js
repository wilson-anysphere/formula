import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

test("batch CLI can migrate + validate a directory of workbook fixtures", async () => {
  const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const cliPath = path.join(pkgRoot, "src", "cli.js");
  const fixturesDir = path.join(pkgRoot, "test", "fixtures", "batch");

  const { stdout } = await execFileAsync("node", [cliPath, "--dir", fixturesDir, "--target", "python"], {
    cwd: pkgRoot,
    maxBuffer: 10 * 1024 * 1024,
  });

  const summary = JSON.parse(stdout.toString("utf8"));
  assert.equal(summary.ok, true);
  assert.equal(summary.totals.files, 1);
  assert.equal(summary.totals.failed, 0);
  assert.equal(summary.results[0].results[0].target, "python");
  assert.equal(summary.results[0].results[0].ok, true);
});

