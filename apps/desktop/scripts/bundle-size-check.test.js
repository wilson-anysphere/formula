import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..");
const scriptPath = path.join(repoRoot, "apps", "desktop", "scripts", "bundle-size-check.mjs");

/**
 * @param {string} distDir
 * @param {{ env?: Record<string, string> }} [opts]
 */
function run(distDir, opts = {}) {
  const env = { ...process.env, ...(opts.env ?? {}) };
  const proc = spawnSync(process.execPath, [scriptPath, "--dist", distDir], {
    cwd: repoRoot,
    encoding: "utf8",
    env,
  });
  if (proc.error) throw proc.error;
  return proc;
}

/**
 * @param {{ files: Record<string, string>, indexHtml: string }} fixture
 */
function writeFixture({ files, indexHtml }) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "formula-desktop-bundle-size-"));
  const distDir = path.join(root, "dist");
  fs.mkdirSync(distDir, { recursive: true });
  fs.writeFileSync(path.join(distDir, "index.html"), indexHtml, "utf8");
  for (const [relPath, content] of Object.entries(files)) {
    const absPath = path.join(distDir, relPath);
    fs.mkdirSync(path.dirname(absPath), { recursive: true });
    fs.writeFileSync(absPath, content, "utf8");
  }
  return { root, distDir };
}

test("reports sizes and exits 0 with no budgets", () => {
  const { root, distDir } = writeFixture({
    indexHtml: `<!doctype html><html><head><script type="module" src="/assets/entry.js"></script><link rel="modulepreload" href="/assets/preload.js"></head></html>`,
    files: {
      "assets/entry.js": "console.log('entry');\n",
      "assets/preload.js": "console.log('preload');\n",
      "coi-check-worker.js": "console.log('worker');\n",
    },
  });

  const proc = run(distDir);
  fs.rmSync(root, { recursive: true, force: true });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Desktop JS bundle size/i);
  assert.match(proc.stdout, /Total JS \(dist\/assets\/\*\*\/\*\.js\)/);
  assert.match(proc.stdout, /Total JS \(dist\/\*\*\/\*\.js\)/);
  assert.match(proc.stdout, /Entry JS \(script tags\)/);
  assert.match(proc.stdout, /assets\/entry\.js/);
});

test("fails when index.html has no entry script tags", () => {
  const { root, distDir } = writeFixture({
    indexHtml: `<!doctype html><html><head></head><body>No scripts</body></html>`,
    files: {
      "assets/entry.js": "console.log('unused');\n",
    },
  });

  const proc = run(distDir);
  fs.rmSync(root, { recursive: true, force: true });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /No JS <script src/);
});

test("fails when budgets are exceeded", () => {
  const { root, distDir } = writeFixture({
    indexHtml: `<!doctype html><html><head><script type="module" src="/assets/entry.js"></script></head></html>`,
    files: {
      "assets/entry.js": "x".repeat(2048), // 2 KiB
    },
  });

  const proc = run(distDir, {
    env: {
      FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB: "1",
      FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB: "1",
    },
  });
  fs.rmSync(root, { recursive: true, force: true });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /JS bundle size budgets exceeded/i);
  assert.match(proc.stderr, /FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB/);
  assert.match(proc.stderr, /FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB/);
});

test("warn-only prints violations but exits 0", () => {
  const { root, distDir } = writeFixture({
    indexHtml: `<!doctype html><html><head><script type="module" src="/assets/entry.js"></script></head></html>`,
    files: {
      "assets/entry.js": "x".repeat(2048), // 2 KiB
    },
  });

  const proc = run(distDir, {
    env: {
      FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB: "1",
      FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB: "1",
      FORMULA_DESKTOP_BUNDLE_SIZE_WARN_ONLY: "1",
    },
  });
  fs.rmSync(root, { recursive: true, force: true });

  assert.equal(proc.status, 0);
  assert.match(proc.stderr, /JS bundle size budgets exceeded/i);
});

test("optional dist total budget (dist/**/*.js) can be enforced separately", () => {
  const { root, distDir } = writeFixture({
    indexHtml: `<!doctype html><html><head><script type="module" src="/assets/entry.js"></script></head></html>`,
    files: {
      "assets/entry.js": "x".repeat(512), // 0.5 KiB
      "coi-check-worker.js": "y".repeat(2048), // 2 KiB (outside assets)
    },
  });

  const proc = run(distDir, {
    env: {
      // Vite bundles are under budget...
      FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB: "10",
      FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB: "10",
      // ...but total JS across dist should trip.
      FORMULA_DESKTOP_JS_DIST_TOTAL_BUDGET_KB: "1",
    },
  });
  fs.rmSync(root, { recursive: true, force: true });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /FORMULA_DESKTOP_JS_DIST_TOTAL_BUDGET_KB/);
});
