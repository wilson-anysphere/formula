import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-pyodide-cache.py");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const pythonProbe = spawnSync("python3", ["--version"], { encoding: "utf8" });
const hasPython3 = !pythonProbe.error && pythonProbe.status === 0;

const canRun = hasPython3;

/**
 * Runs the pyodide cache guard in a temporary directory that mimics the repo layout.
 * @param {{ ensureScript: string; workflowYaml: string }} input
 */
function run({ ensureScript, workflowYaml }) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-pyodide-cache-"));
  try {
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });
    mkdirSync(path.join(tmpdir, "apps", "desktop", "scripts"), { recursive: true });
    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });

    writeFileSync(path.join(tmpdir, "scripts", "ci", "check-pyodide-cache.py"), scriptContents, "utf8");
    writeFileSync(
      path.join(tmpdir, "apps", "desktop", "scripts", "ensure-pyodide-assets.mjs"),
      `${ensureScript}\n`,
      "utf8",
    );
    writeFileSync(path.join(tmpdir, ".github", "workflows", "workflow.yml"), `${workflowYaml}\n`, "utf8");

    return spawnSync("python3", ["scripts/ci/check-pyodide-cache.py"], { cwd: tmpdir, encoding: "utf8" });
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("passes when desktop build job includes pyodide cache steps + key", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `// minimal ensure script\nconst PYODIDE_VERSION = \"0.26.4\";\nexport {};`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo \"version=0.26.4\" >> \"$GITHUB_OUTPUT\"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Pyodide cache guard: OK/i);
});

test("passes when pyodide cache key is quoted", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = '0.26.4';`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo "version=0.26.4" >> "$GITHUB_OUTPUT"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/
          key: "pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}"

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/
          key: "pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}"

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Pyodide cache guard: OK/i);
});

test("repo workflows satisfy pyodide cache guardrails", { skip: !canRun }, () => {
  // Defense-in-depth: ensure our real workflows keep the (gated) Pyodide cache steps in place so
  // turning on `FORMULA_BUNDLE_PYODIDE_ASSETS=1` doesn't regress into repeated downloads.
  const proc = spawnSync("python3", ["scripts/ci/check-pyodide-cache.py"], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Pyodide cache guard: OK/i);
});

test("fails when desktop build job is missing pyodide cache key", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = '0.26.4';`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing `key: pyodide/i);
});

test("passes when caching the whole pyodide tree restores tracked files and enables cross-OS fallback", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = "0.26.4";`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo "version=0.26.4" >> "$GITHUB_OUTPUT"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          enableCrossOsArchive: true
          path: apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}
          restore-keys: |
            pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-
            pyodide-\${{ steps.pyodide.outputs.version }}-

      - name: Restore tracked Pyodide files
        run: git restore --source=HEAD -- apps/desktop/public/pyodide

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          enableCrossOsArchive: true
          path: apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Pyodide cache guard: OK/i);
});

test("fails when caching the whole pyodide tree uses multiline path without restoring tracked files", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = "0.26.4";`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo "version=0.26.4" >> "$GITHUB_OUTPUT"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          enableCrossOsArchive: true
          path: |
            apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}
          restore-keys: |
            pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-
            pyodide-\${{ steps.pyodide.outputs.version }}-

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          enableCrossOsArchive: true
          path: |
            apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /does not restore tracked files/i);
});

test("fails when caching the whole pyodide tree does not restore tracked files", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = "0.26.4";`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo "version=0.26.4" >> "$GITHUB_OUTPUT"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with:
          enableCrossOsArchive: true
          path: apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}
          restore-keys: |
            pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-
            pyodide-\${{ steps.pyodide.outputs.version }}-

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with:
          enableCrossOsArchive: true
          path: apps/desktop/public/pyodide/
          key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }}

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /does not restore tracked files/i);
});

test("passes when pyodide cache restore/save use inline flow mapping syntax", { skip: !canRun }, () => {
  const proc = run({
    ensureScript: `const PYODIDE_VERSION = "0.26.4";`,
    workflowYaml: `
name: Desktop build
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Detect Pyodide version (for caching)
        id: pyodide
        run: echo "version=0.26.4" >> "$GITHUB_OUTPUT"

      - name: Restore Pyodide asset cache
        uses: actions/cache/restore@v4
        with: { path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/, key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }} }

      - name: Ensure Pyodide assets are present (populate cache on miss)
        run: node apps/desktop/scripts/ensure-pyodide-assets.mjs

      - name: Save Pyodide asset cache
        uses: actions/cache/save@v4
        with: { path: apps/desktop/public/pyodide/v\${{ steps.pyodide.outputs.version }}/full/, key: pyodide-\${{ runner.os }}-\${{ steps.pyodide.outputs.version }}-\${{ hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs') }} }

      - name: Build desktop frontend assets
        run: pnpm build:desktop
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Pyodide cache guard: OK/i);
});
