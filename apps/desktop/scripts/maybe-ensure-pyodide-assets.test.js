import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..");
const scriptSourcePath = path.join(repoRoot, "apps", "desktop", "scripts", "maybe-ensure-pyodide-assets.mjs");
const scriptSource = fs.readFileSync(scriptSourcePath, "utf8");

/**
 * @param {{
 *   env?: Record<string, string | undefined>,
 *   publicFiles?: Record<string, string>
 * }} opts
 */
function runFixture(opts = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "formula-maybe-pyodide-"));
  const appsDesktop = path.join(root, "apps", "desktop");
  const scriptsDir = path.join(appsDesktop, "scripts");
  const publicDir = path.join(appsDesktop, "public");

  fs.mkdirSync(scriptsDir, { recursive: true });
  fs.mkdirSync(publicDir, { recursive: true });

  // Write the real script under test.
  const maybePath = path.join(scriptsDir, "maybe-ensure-pyodide-assets.mjs");
  fs.writeFileSync(maybePath, scriptSource, "utf8");

  // Stub ensure script so tests never hit the network.
  const ensurePath = path.join(scriptsDir, "ensure-pyodide-assets.mjs");
  fs.writeFileSync(
    ensurePath,
    [
      "import { mkdir, writeFile } from 'node:fs/promises';",
      "import path from 'node:path';",
      "import { fileURLToPath } from 'node:url';",
      "const __dirname = path.dirname(fileURLToPath(import.meta.url));",
      "const out = path.resolve(__dirname, '../public/pyodide/STUB_ASSET');",
      "await mkdir(path.dirname(out), { recursive: true });",
      "await writeFile(out, 'stub', 'utf8');",
      "export {};",
      "",
    ].join("\n"),
    "utf8",
  );

  for (const [relPath, contents] of Object.entries(opts.publicFiles ?? {})) {
    const absPath = path.join(appsDesktop, "public", relPath);
    fs.mkdirSync(path.dirname(absPath), { recursive: true });
    fs.writeFileSync(absPath, contents, "utf8");
  }

  const env = { ...process.env };
  for (const [k, v] of Object.entries(opts.env ?? {})) {
    if (v == null) delete env[k];
    else env[k] = v;
  }

  const proc = spawnSync(process.execPath, [maybePath], {
    cwd: root,
    env,
    encoding: "utf8",
  });

  return { root, appsDesktop, proc };
}

test("removes public/pyodide when bundling flag is not set", () => {
  const { root, appsDesktop, proc } = runFixture({
    publicFiles: {
      "pyodide/v0.25.1/full/python_stdlib.zip": "stub",
      "pyodide/other.txt": "stub",
    },
    env: { FORMULA_BUNDLE_PYODIDE_ASSETS: undefined },
  });

  try {
    assert.equal(proc.status, 0, proc.stderr);
    assert.equal(
      fs.existsSync(path.join(appsDesktop, "public", "pyodide")),
      false,
      "expected public/pyodide to be removed",
    );
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("runs ensure script when bundling flag is enabled", () => {
  const { root, appsDesktop, proc } = runFixture({
    publicFiles: {
      "pyodide/v0.25.1/full/python_stdlib.zip": "stub",
    },
    env: { FORMULA_BUNDLE_PYODIDE_ASSETS: "1" },
  });

  try {
    assert.equal(proc.status, 0, proc.stderr);
    assert.equal(
      fs.readFileSync(path.join(appsDesktop, "public", "pyodide", "STUB_ASSET"), "utf8"),
      "stub",
      "expected ensure script to run",
    );
    assert.equal(
      fs.readFileSync(path.join(appsDesktop, "public", "pyodide", "v0.25.1", "full", "python_stdlib.zip"), "utf8"),
      "stub",
      "expected existing assets to remain when bundling is enabled",
    );
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

