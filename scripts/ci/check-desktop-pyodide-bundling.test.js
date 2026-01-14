import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

function readJson(relPath) {
  const absPath = path.join(repoRoot, relPath);
  return JSON.parse(readFileSync(absPath, "utf8"));
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
