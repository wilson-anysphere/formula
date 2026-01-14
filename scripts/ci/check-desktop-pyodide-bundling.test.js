import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments, stripRustComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

function readJson(relPath) {
  const absPath = path.join(repoRoot, relPath);
  return JSON.parse(readFileSync(absPath, "utf8"));
}

function readText(relPath) {
  const absPath = path.join(repoRoot, relPath);
  return readFileSync(absPath, "utf8");
}

function extractPyodideVersionFromEnsureScript(src) {
  const m = src.match(/const\s+PYODIDE_VERSION\s*=\s*['"]([^'"]+)['"]/);
  assert.ok(m, "Expected ensure-pyodide-assets.mjs to define const PYODIDE_VERSION");
  return m[1];
}

function extractPyodideVersionFromRust(src) {
  const m = src.match(/const\s+PYODIDE_VERSION\s*:\s*&str\s*=\s*"([^"]+)"/);
  assert.ok(m, "Expected pyodide_assets.rs to define const PYODIDE_VERSION: &str");
  return m[1];
}

function extractPyodideVersionFromCdnUrl(src, context) {
  const m = src.match(/https:\/\/cdn\.jsdelivr\.net\/pyodide\/v([0-9]+\.[0-9]+\.[0-9]+)\/full\//);
  assert.ok(m, `Expected ${context} to reference a cdn.jsdelivr.net/pyodide/vX.Y.Z/full/ URL`);
  return m[1];
}

function extractPyodideVersionFromLocalPyodideUrl(src, context) {
  const m = src.match(/\/pyodide\/v([0-9]+\.[0-9]+\.[0-9]+)\/full\//);
  assert.ok(m, `Expected ${context} to reference a /pyodide/vX.Y.Z/full/ URL`);
  return m[1];
}

function extractRequiredFilesFromEnsureScript(src) {
  const m = src.match(/const\s+requiredFiles\s*=\s*\{([\s\S]*?)\n\};/m);
  assert.ok(m, "Expected ensure-pyodide-assets.mjs to define `const requiredFiles = { ... };`");
  const body = m[1];

  const entries = new Map();
  const re = /['"]([^'"]+)['"]\s*:\s*['"]([0-9a-f]{64})['"]/g;
  for (const match of body.matchAll(re)) {
    entries.set(match[1], match[2]);
  }
  assert.ok(entries.size > 0, "Expected ensure-pyodide-assets.mjs requiredFiles to be non-empty");
  return entries;
}

function extractRequiredFilesFromRust(src) {
  const blockMatch = src.match(/const\s+PYODIDE_REQUIRED_FILES[\s\S]*?=\s*&\[\s*([\s\S]*?)\s*\];/m);
  assert.ok(blockMatch, "Expected pyodide_assets.rs to define PYODIDE_REQUIRED_FILES");
  const block = blockMatch[1];

  const entries = new Map();
  const re = /file_name:\s*"([^"]+)"[\s\S]*?sha256:\s*"([0-9a-f]{64})"/g;
  for (const match of block.matchAll(re)) {
    entries.set(match[1], match[2]);
  }
  assert.ok(entries.size > 0, "Expected pyodide_assets.rs PYODIDE_REQUIRED_FILES to be non-empty");
  return entries;
}

test("desktop dev/build scripts do not bundle Pyodide by default", () => {
  const pkg = readJson("apps/desktop/package.json");
  const scripts = pkg?.scripts ?? {};

  assert.equal(typeof scripts.dev, "string");
  assert.equal(typeof scripts.build, "string");

  // The large Pyodide distribution should not be downloaded/copied into `dist/` on every build.
  // The desktop app now downloads Pyodide on-demand at runtime and caches it in the app data dir.
  assert.match(
    scripts.dev,
    /\bmaybe-ensure-pyodide-assets\.mjs\b/,
    "Expected apps/desktop/package.json#scripts.dev to invoke maybe-ensure-pyodide-assets.mjs",
  );
  assert.match(
    scripts.build,
    /\bmaybe-ensure-pyodide-assets\.mjs\b/,
    "Expected apps/desktop/package.json#scripts.build to invoke maybe-ensure-pyodide-assets.mjs",
  );

  // Defense-in-depth: ensure we do not regress to always running the downloader.
  assert.ok(
    !scripts.dev.includes("scripts/ensure-pyodide-assets.mjs"),
    "Expected apps/desktop/package.json#scripts.dev to not call ensure-pyodide-assets.mjs directly (should be opt-in)",
  );
  assert.ok(
    !scripts.build.includes("scripts/ensure-pyodide-assets.mjs"),
    "Expected apps/desktop/package.json#scripts.build to not call ensure-pyodide-assets.mjs directly (should be opt-in)",
  );
});

test("maybe-ensure script gates bundling behind FORMULA_BUNDLE_PYODIDE_ASSETS", () => {
  const absPath = path.join(repoRoot, "apps/desktop/scripts/maybe-ensure-pyodide-assets.mjs");
  const src = stripComments(readFileSync(absPath, "utf8"));

  assert.match(
    src,
    /FORMULA_BUNDLE_PYODIDE_ASSETS/,
    "Expected maybe-ensure-pyodide-assets.mjs to reference FORMULA_BUNDLE_PYODIDE_ASSETS",
  );
  assert.match(
    src,
    /ensure-pyodide-assets\.mjs/,
    "Expected maybe-ensure-pyodide-assets.mjs to be able to invoke ensure-pyodide-assets.mjs when opted in",
  );
});

test("vite build strips dist/pyodide unless FORMULA_BUNDLE_PYODIDE_ASSETS is set", () => {
  const src = stripComments(readText("apps/desktop/vite.config.ts"));

  // Defense-in-depth: even if a dev/CI cache accidentally populates `apps/desktop/public/pyodide/...`,
  // ensure the production Vite build removes it from `dist/` unless bundling is explicitly enabled.
  assert.match(
    src,
    /stripPyodideFromDist/,
    "Expected apps/desktop/vite.config.ts to include the stripPyodideFromDist build hook",
  );
  assert.match(
    src,
    /FORMULA_BUNDLE_PYODIDE_ASSETS/,
    "Expected apps/desktop/vite.config.ts to gate Pyodide bundling behind FORMULA_BUNDLE_PYODIDE_ASSETS",
  );
  assert.match(
    src,
    /rmSync\([^)]*["']pyodide["']/,
    "Expected apps/desktop/vite.config.ts to remove the pyodide directory from dist via rmSync(...)",
  );
});

test("desktop Pyodide version + required assets stay in sync (ensure script / Rust / python-runtime)", () => {
  const ensureSrc = readText("apps/desktop/scripts/ensure-pyodide-assets.mjs");
  const rustSrc = stripRustComments(readText("apps/desktop/src-tauri/src/pyodide_assets.rs"));

  const ensureVersion = extractPyodideVersionFromEnsureScript(ensureSrc);
  const rustVersion = extractPyodideVersionFromRust(rustSrc);
  assert.equal(
    rustVersion,
    ensureVersion,
    "Expected apps/desktop/src-tauri/src/pyodide_assets.rs PYODIDE_VERSION to match apps/desktop/scripts/ensure-pyodide-assets.mjs",
  );

  const pythonMainThreadSrc = readText("packages/python-runtime/src/pyodide-main-thread.js");
  const pythonWorkerSrc = readText("packages/python-runtime/src/pyodide-worker.js");
  const pythonMainThreadVersion = extractPyodideVersionFromCdnUrl(
    pythonMainThreadSrc,
    "packages/python-runtime/src/pyodide-main-thread.js",
  );
  const pythonWorkerVersion = extractPyodideVersionFromCdnUrl(
    pythonWorkerSrc,
    "packages/python-runtime/src/pyodide-worker.js",
  );
  assert.equal(
    pythonMainThreadVersion,
    ensureVersion,
    "Expected python-runtime main-thread default CDN index URL to match desktop PYODIDE_VERSION",
  );
  assert.equal(
    pythonWorkerVersion,
    ensureVersion,
    "Expected python-runtime worker default CDN index URL to match desktop PYODIDE_VERSION",
  );

  // E2E harness uses a self-hosted `/pyodide/...` URL; keep it aligned so tests don't silently
  // pin a different version than the runtime downloader.
  const e2eHtml = readText("apps/desktop/python-runtime-test.html");
  const e2eVersion = extractPyodideVersionFromLocalPyodideUrl(
    e2eHtml,
    "apps/desktop/python-runtime-test.html",
  );
  assert.equal(e2eVersion, ensureVersion, "Expected python-runtime-test.html to match desktop PYODIDE_VERSION");

  const ensureFiles = extractRequiredFilesFromEnsureScript(ensureSrc);
  const rustFiles = extractRequiredFilesFromRust(rustSrc);

  assert.equal(
    rustFiles.size,
    ensureFiles.size,
    "Expected Rust PYODIDE_REQUIRED_FILES to match ensure-pyodide-assets.mjs requiredFiles entry count",
  );

  for (const [fileName, sha] of ensureFiles.entries()) {
    assert.equal(
      rustFiles.get(fileName),
      sha,
      `Expected Rust sha256 for ${fileName} to match ensure-pyodide-assets.mjs`,
    );
  }
});
