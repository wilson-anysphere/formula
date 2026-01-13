import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, truncateSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import process from "node:process";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "desktop_dist_asset_report.mjs");

/**
 * @param {string} filePath
 * @param {number} sizeBytes
 */
function createSizedFile(filePath, sizeBytes) {
  mkdirSync(path.dirname(filePath), { recursive: true });
  // Ensure the file exists before truncating.
  writeFileSync(filePath, "", "utf8");
  truncateSync(filePath, sizeBytes);
}

test("desktop_dist_asset_report emits markdown with top offenders + grouped totals", () => {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-desktop-dist-report-"));
  createSizedFile(path.join(distDir, "assets", "a.bin"), 2_000_000);
  createSizedFile(path.join(distDir, "pyodide", "b.wasm"), 5_000_000);

  const proc = spawnSync(process.execPath, [scriptPath, "--dist-dir", distDir, "--top", "2", "--group-depth", "2"], {
    encoding: "utf8",
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /## Desktop dist asset report/);
  assert.match(proc.stdout, /### Top 2 largest files/);
  assert.match(proc.stdout, /`pyodide\/b\.wasm`/);
  assert.match(proc.stdout, /### Grouped totals/);
  assert.match(proc.stdout, /`pyodide\/`/);
  assert.match(proc.stdout, /`assets\/`/);
  assert.doesNotMatch(proc.stdout, /`assets\/a\.bin\/`/);
});

test("desktop_dist_asset_report tolerates pnpm-style -- delimiter", () => {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-desktop-dist-delimiter-"));
  createSizedFile(path.join(distDir, "assets", "a.bin"), 123);
  const proc = spawnSync(process.execPath, [scriptPath, "--", "--dist-dir", distDir, "--top", "1", "--no-groups"], {
    encoding: "utf8",
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /## Desktop dist asset report/);
  assert.doesNotMatch(proc.stdout, /### Grouped totals/);
});

test("desktop_dist_asset_report resolves missing relative --dist-dir against CWD (not repo root)", () => {
  const cwd = mkdtempSync(path.join(tmpdir(), "formula-desktop-dist-cwd-"));
  const missingName = `missing-dist-${process.pid}-${Date.now()}`;
  const expected = path.join(cwd, missingName);

  const proc = spawnSync(process.execPath, [scriptPath, "--dist-dir", missingName], {
    encoding: "utf8",
    cwd,
  });

  assert.equal(proc.status, 1);
  assert.match(proc.stderr, /dist directory not found/i);
  assert.ok(proc.stdout.includes(`Dist dir: \`${expected}\``), proc.stdout);
});

test("desktop_dist_asset_report enforces budgets when env vars are set", () => {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-desktop-dist-budget-"));
  createSizedFile(path.join(distDir, "assets", "a.bin"), 2_000_000);
  createSizedFile(path.join(distDir, "pyodide", "b.wasm"), 5_000_000);

  const proc = spawnSync(process.execPath, [scriptPath, "--dist-dir", distDir, "--top", "1"], {
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB: "6",
      FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB: "3",
    },
  });

  assert.equal(proc.status, 1);
  assert.match(proc.stderr, /exceed single-file budget/i);
  assert.match(proc.stderr, /FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB/);
  assert.match(proc.stdout, /Budgets:/);
  assert.match(proc.stdout, /\*\*FAIL\*\*/);
});
