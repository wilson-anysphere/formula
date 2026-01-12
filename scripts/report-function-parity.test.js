import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = resolve(repoRoot, "scripts", "report-function-parity.mjs");

function runScript() {
  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return proc.stdout;
}

function extractSection(lines, headingRe) {
  const matches = [];
  for (let i = 0; i < lines.length; i++) {
    if (headingRe.test(lines[i])) matches.push(i);
  }
  assert.ok(matches.length > 0, `expected section heading ${headingRe}`);
  const start = matches[matches.length - 1] + 1;

  /** @type {string[]} */
  const names = [];
  for (let i = start; i < lines.length; i++) {
    const line = lines[i];
    if (line.trim().length === 0) break;
    const m = line.match(/^  - (.+)$/);
    if (!m) continue;
    names.push(m[1]);
  }
  return names;
}

test("report-function-parity script runs and produces deterministic sorted lists", () => {
  const out = runScript();
  const lines = out.split("\n").map((line) => line.trimEnd());

  assert.ok(lines.some((line) => line.startsWith("Catalog functions (shared/functionCatalog.json): ")));
  assert.ok(lines.some((line) => line.startsWith("FTAB functions (crates/formula-biff/src/ftab.rs): ")));

  const missing = extractSection(lines, /^FTAB \\ Catalog \(missing from catalog\): \d+$/);
  const notInFtab = extractSection(lines, /^Catalog \\ FTAB \(not present in FTAB\): \d+$/);

  // The script prints "top N" lists (currently N=50).
  assert.ok(missing.length <= 50, "missing-from-catalog list should be limited to TOP_N");
  assert.ok(notInFtab.length <= 50, "catalog-not-in-ftab list should be limited to TOP_N");

  const missingSorted = [...missing].sort();
  assert.deepEqual(missing, missingSorted, "missing-from-catalog list should be sorted");
  assert.equal(missing.length, new Set(missing).size, "missing-from-catalog list should be unique");

  const notInFtabSorted = [...notInFtab].sort();
  assert.deepEqual(notInFtab, notInFtabSorted, "catalog-not-in-ftab list should be sorted");
  assert.equal(notInFtab.length, new Set(notInFtab).size, "catalog-not-in-ftab list should be unique");
});

