import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import process from "node:process";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "frontend_asset_size_report.mjs");

function createFixtureDistDir() {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-frontend-asset-size-"));
  const assetsDir = path.join(distDir, "assets");
  mkdirSync(assetsDir, { recursive: true });

  writeFileSync(path.join(assetsDir, "app.js"), "console.log('hello');\n".repeat(2000), "utf8");
  writeFileSync(path.join(assetsDir, "app.js.map"), "{}", "utf8");
  writeFileSync(path.join(assetsDir, "styles.css"), "body{color:red;}\n".repeat(2000), "utf8");
  writeFileSync(path.join(assetsDir, "engine.wasm"), Buffer.alloc(1024, 0));

  return distDir;
}

test("frontend_asset_size_report emits markdown and ignores sourcemaps", () => {
  const distDir = createFixtureDistDir();

  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir], {
    encoding: "utf8",
    env: {
      ...process.env,
      // GitHub Actions `vars.*` interpolate to empty strings when unset; ensure this is handled.
      FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION: "",
    },
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /## Frontend asset download size/);
  assert.match(proc.stdout, /Totals: raw/);
  assert.match(proc.stdout, /`.*app\.js`/);
  assert.match(proc.stdout, /`.*styles\.css`/);
  assert.match(proc.stdout, /`.*engine\.wasm`/);
  assert.doesNotMatch(proc.stdout, /\.map`/);
});

test("frontend_asset_size_report enforces budget when enabled", () => {
  const distDir = createFixtureDistDir();

  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir, "--enforce", "--limit-mb", "0.000001"], {
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION: "gzip",
    },
  });

  assert.equal(proc.status, 1);
  assert.match(proc.stderr, /frontend-asset-size: ERROR/i);
  assert.match(proc.stderr, /exceeds/i);
});

test("frontend_asset_size_report returns a markdown report when dist/assets is missing", () => {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-frontend-asset-missing-"));

  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir], {
    encoding: "utf8",
  });

  assert.equal(proc.status, 1);
  assert.match(proc.stderr, /missing Vite assets directory/i);
  assert.match(proc.stdout, /## Frontend asset download size/);
  assert.match(proc.stdout, /directory not found/i);
});

test("frontend_asset_size_report writes a JSON report when --json-out is provided", () => {
  const distDir = createFixtureDistDir();
  const jsonOut = path.join(distDir, "frontend-asset-size.json");

  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir, "--json-out", jsonOut], {
    encoding: "utf8",
    env: {
      ...process.env,
      // GitHub Actions `vars.*` interpolate to empty strings when unset; ensure this is handled.
      FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION: "",
    },
  });

  assert.equal(proc.status, 0, proc.stderr);
  const report = JSON.parse(readFileSync(jsonOut, "utf8"));
  assert.equal(report.missing_assets_dir, false);
  assert.equal(report.file_count, 3);
  assert.equal(report.compression, "brotli");
  assert.equal(report.limit_mb, 10);
  assert.equal(report.enforce, false);
  assert.equal(report.over_limit, false);
  assert.ok(report.totals?.raw_bytes > 0);
  assert.ok(report.totals?.brotli_bytes > 0);
  assert.ok(report.totals?.gzip_bytes > 0);
  assert.ok(report.assets.some((a) => String(a.path).includes("app.js")));
  assert.ok(report.assets.some((a) => String(a.path).includes("styles.css")));
  assert.ok(report.assets.some((a) => String(a.path).includes("engine.wasm")));
  assert.equal(report.assets.some((a) => String(a.path).endsWith(".map")), false);
});

test("frontend_asset_size_report writes a JSON report even when dist/assets is missing", () => {
  const distDir = mkdtempSync(path.join(tmpdir(), "formula-frontend-asset-missing-json-"));
  const jsonOut = path.join(distDir, "frontend-asset-size.json");

  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir, "--json-out", jsonOut], {
    encoding: "utf8",
  });

  assert.equal(proc.status, 1);
  const report = JSON.parse(readFileSync(jsonOut, "utf8"));
  assert.equal(report.missing_assets_dir, true);
  assert.equal(report.file_count, 0);
  assert.equal(Array.isArray(report.assets), true);
  assert.equal(report.assets.length, 0);
});

test("frontend_asset_size_report writes a JSON report even when enforcing failures occur", () => {
  const distDir = createFixtureDistDir();
  const jsonOut = path.join(distDir, "frontend-asset-size.json");

  const proc = spawnSync(
    process.execPath,
    [scriptPath, "--dist", distDir, "--enforce", "--limit-mb", "0.000001", "--compression", "gzip", "--json-out", jsonOut],
    {
      encoding: "utf8",
    },
  );

  assert.equal(proc.status, 1);
  const report = JSON.parse(readFileSync(jsonOut, "utf8"));
  assert.equal(report.missing_assets_dir, false);
  assert.equal(report.enforce, true);
  assert.equal(report.over_limit, true);
  assert.equal(report.compression, "gzip");
  assert.ok(report.totals?.compressed_bytes > 0);
});
