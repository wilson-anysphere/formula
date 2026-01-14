import { spawn, spawnSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import { readdir, readFile, stat } from "node:fs/promises";
import { builtinModules, createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { stripComments } from "../test/sourceTextUtils.js";

/**
 * Node's `--test` runner started detecting TypeScript test files (`*.test.ts`) once
 * TypeScript stripping support landed. `apps/desktop` uses Vitest for `.test.ts`
 * suites, while `test:node` is intended to run only `node:test` suites written in
 * JavaScript.
 *
 * Run an explicit list of test files so `pnpm -C apps/desktop test:node` stays
 * stable across Node.js versions.
 */

const desktopRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(desktopRoot, "../..");
const require = createRequire(import.meta.url);

/** @type {string[]} */
const files = [];
await collectTests(desktopRoot, files);
files.sort((a, b) => a.localeCompare(b));

const tsLoaderArgs = resolveTypeScriptLoaderArgs();
const builtInTypeScript = getBuiltInTypeScriptSupport();
const canExecuteTypeScript = tsLoaderArgs.length > 0 || builtInTypeScript.enabled;

// Node's built-in "strip types" support can execute `.ts` modules, but does not support
// `.tsx` (JSX) without a real transpile loader.
const canExecuteTsx = tsLoaderArgs.length > 0;

let runnableFiles = files;
let typeScriptFilteredCount = 0;
let typeScriptTsxFilteredCount = 0;
if (!canExecuteTypeScript) {
  runnableFiles = await filterTypeScriptImportTests(files, ["ts", "tsx"]);
  typeScriptFilteredCount = files.length - runnableFiles.length;
} else if (!canExecuteTsx) {
  runnableFiles = await filterTypeScriptImportTests(files, ["tsx"]);
  typeScriptTsxFilteredCount = files.length - runnableFiles.length;
}

const hasDeps = await hasNodeModules();
let externalDepsFilteredCount = 0;
let missingWorkspaceDepsFilteredCount = 0;
if (!hasDeps) {
  const before = runnableFiles.length;
  runnableFiles = await filterExternalDependencyTests(runnableFiles, {
    canStripTypes: canExecuteTypeScript,
    canExecuteTsx,
  });
  externalDepsFilteredCount = before - runnableFiles.length;
} else {
  const before = runnableFiles.length;
  runnableFiles = await filterMissingWorkspaceDependencyTests(runnableFiles, {
    canStripTypes: canExecuteTypeScript,
    canExecuteTsx,
  });
  missingWorkspaceDepsFilteredCount = before - runnableFiles.length;
}

if (runnableFiles.length !== files.length) {
  const skipped = files.length - runnableFiles.length;
  /** @type {string[]} */
  const reasons = [];
  if (typeScriptFilteredCount > 0) {
    reasons.push(`${typeScriptFilteredCount} import TypeScript modules (TypeScript execution not available)`);
  }
  if (typeScriptTsxFilteredCount > 0) {
    reasons.push(`${typeScriptTsxFilteredCount} import .tsx modules (TSX execution not available)`);
  }
  if (externalDepsFilteredCount > 0) {
    reasons.push(`${externalDepsFilteredCount} depend on external packages (dependencies not installed)`);
  }
  if (missingWorkspaceDepsFilteredCount > 0) {
    reasons.push(`${missingWorkspaceDepsFilteredCount} depend on missing workspace packages`);
  }
  const suffix = reasons.length > 0 ? ` (${reasons.join("; ")})` : "";
  console.log(`Skipping ${skipped} node:test file(s) that can't run in this environment${suffix}.`);
}

if (runnableFiles.length === 0) {
  if (files.length === 0) {
    console.log("No node:test files found.");
  } else {
    console.log("All node:test files were skipped.");
  }
  process.exit(0);
}

const baseNodeArgs = ["--no-warnings"];
if (tsLoaderArgs.length > 0) {
  baseNodeArgs.push(...tsLoaderArgs);
} else if (builtInTypeScript.enabled) {
  baseNodeArgs.push(...builtInTypeScript.args);
  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
  baseNodeArgs.push(...resolveNodeLoaderArgs(loaderUrl));
}

// Node's test runner defaults to process isolation, which can spawn many Node processes.
// In extremely constrained environments (like some agent sandboxes) we can hit OS thread
// limits due to the ESM loader hook worker that Node creates for each isolated process.
//
// When third-party deps are missing (`hasDeps === false`), these node:test runs are already
// best-effort guardrails. Prefer running without isolation in that mode so the suite is
// more resilient to low thread limits.
//
// Allow overriding via `FORMULA_NODE_TEST_ISOLATION` (e.g. "process", "none").
const requestedIsolation = typeof process.env.FORMULA_NODE_TEST_ISOLATION === "string" ? process.env.FORMULA_NODE_TEST_ISOLATION.trim() : "";
const allowedFlags =
  process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
    ? process.allowedNodeEnvironmentFlags
    : new Set();
if (allowedFlags.has("--test-isolation")) {
  const isolation = requestedIsolation || (hasDeps ? "" : "none");
  if (isolation) baseNodeArgs.push(`--test-isolation=${isolation}`);
}

// Keep node:test parallelism conservative; some suites start background services and
// in CI/agent environments we can hit process/thread limits if too many test files
// run in parallel. Allow opting into higher parallelism via FORMULA_NODE_TEST_CONCURRENCY.
const parsedConcurrency = Number.parseInt(process.env.FORMULA_NODE_TEST_CONCURRENCY ?? "", 10);
const concurrency = Number.isFinite(parsedConcurrency) && parsedConcurrency > 0 ? parsedConcurrency : 1;
const supportsTestConcurrency = supportsNodeFlag(`--test-concurrency=${concurrency}`);
const nodeArgs = [...baseNodeArgs];
if (supportsTestConcurrency) nodeArgs.push(`--test-concurrency=${concurrency}`);
nodeArgs.push("--test", ...runnableFiles);
const child = spawn(process.execPath, nodeArgs, { stdio: "inherit" });
child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

/**
 * Some `node --test` flags are relatively new. We want this runner to stay usable even when
 * a sandbox pins an older Node.js version (engines: >=18), so probe for support before
 * passing optional flags.
 *
 * @param {string} flag
 * @returns {boolean}
 */
function supportsNodeFlag(flag) {
  const probe = spawnSync(process.execPath, [flag, "--version"], { stdio: "ignore" });
  return probe.status === 0;
}

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collectTests(dir, out) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    // Keep this list in sync with `scripts/run-node-tests.mjs` at the repo root.
    if (
      entry.name === ".git" ||
      entry.name === "node_modules" ||
      entry.name === "dist" ||
      entry.name === "coverage" ||
      entry.name === "target" ||
      entry.name === "build" ||
      entry.name === ".turbo" ||
      entry.name === ".pnpm-store" ||
      entry.name === ".cache" ||
      entry.name === ".vite" ||
      entry.name === "playwright-report" ||
      entry.name === "test-results" ||
      entry.name === "security-report" ||
      entry.name.startsWith(".tmp")
    ) {
      continue;
    }

    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectTests(fullPath, out);
      continue;
    }

    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".test.js")) continue;
    out.push(fullPath);
  }
}

/**
 * @param {string[]} files
 * @param {("ts" | "tsx")[]} extensions
 */
async function filterTypeScriptImportTests(files, extensions = ["ts", "tsx"]) {
  /** @type {string[]} */
  const out = [];
  const extGroup = extensions.join("|");
  const tsImportRe = new RegExp(
    `from\\s+["'\`][^"'\`]+\\.(${extGroup})["'\`]|import\\(\\s*["'\`][^"'\`]+\\.(${extGroup})["'\`]\\s*(?:\\)|,)`,
  );

  // In environments without TS/TSX execution, we also need to skip tests that import
  // workspace packages whose entrypoints are authored as `.ts`/`.tsx` (even if the test
  // itself doesn't directly import a `.ts`/`.tsx` file).
  const workspacePackages = await loadWorkspacePackageEntrypoints();
  const disallowedEntrypointExtensions = new Set(extensions.map((ext) => `.${ext}`));
  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  // Support `import("pkg")` and `import("pkg", { ... })` (import options / assertions).
  const dynamicImportRe = /\bimport\s*\(\s*["']([^"']+)["']\s*(?:\)|,)/g;
  // Template literal dynamic imports (used for cache-busting query strings, etc.). We capture
  // only the static prefix up to the first unescaped `${...}` so we can still resolve the
  // underlying module path (e.g. `./foo.js?x=${Date.now()}` -> `./foo.js`).
  const dynamicImportTemplateRe = /\bimport\s*\(\s*`((?:\\.|[^`$])*)/g;
  const requireCallRe = /\brequire\(\s*["']([^"']+)["']\s*\)/g;

  /**
   * @param {string} specifier
   * @returns {{ packageName: string, exportKey: string } | null}
   */
  function parseWorkspaceSpecifier(specifier) {
    if (!specifier.startsWith("@formula/")) return null;
    const parts = specifier.split("/");
    if (parts.length < 2) return null;
    const packageName = `${parts[0]}/${parts[1]}`;
    const subpath = parts.slice(2).join("/");
    const exportKey = subpath ? `./${subpath}` : ".";
    return { packageName, exportKey };
  }

  /**
   * @param {any} exportsMap
   * @param {string} exportKey
   * @param {string | null} main
   */
  function resolveExportPath(exportsMap, exportKey, main) {
    /**
     * @param {any} entry
     * @returns {string | null}
     */
    function pickExportPath(entry) {
      if (typeof entry === "string") return entry;
      if (Array.isArray(entry)) {
        for (const item of entry) {
          const picked = pickExportPath(item);
          if (picked) return picked;
        }
        return null;
      }
      if (entry && typeof entry === "object") {
        // Prefer Node's default ESM condition order.
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.require === "string") return entry.require;

        // Fall back to searching nested condition objects.
        for (const value of Object.values(entry)) {
          const picked = pickExportPath(value);
          if (picked) return picked;
        }
      }
      return null;
    }

    let target = null;
    if (exportsMap) {
      if (typeof exportsMap === "string") {
        if (exportKey === ".") target = exportsMap;
      } else if (exportsMap && typeof exportsMap === "object") {
        const keys = Object.keys(exportsMap);
        const looksLikeSubpathMap = keys.some((k) => k === "." || k.startsWith("./"));
        if (looksLikeSubpathMap) {
          target = pickExportPath(exportsMap?.[exportKey]);
        } else if (exportKey === ".") {
          target = pickExportPath(exportsMap);
        }
      }
    }

    // Packages without an `exports` map (like `@formula/marketplace-shared`) still allow deep
    // imports via `@scope/pkg/<path>`. When TS/TSX execution is unavailable, treat those
    // subpaths as the effective entrypoint for skip decisions.
    if (!exportsMap && exportKey !== ".") return exportKey;

    if (!target && exportKey === "." && typeof main === "string") target = main;
    return target;
  }

  /**
   * @param {string} text
   */
  function importsWorkspaceTypeScriptEntrypoint(text) {
    /** @type {string[]} */
    const specifiers = [];
    for (const match of text.matchAll(importFromRe)) specifiers.push(match[2]);
    for (const match of text.matchAll(sideEffectImportRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(dynamicImportRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(dynamicImportTemplateRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(requireCallRe)) specifiers.push(match[1]);

    for (const rawSpecifier of specifiers) {
      if (!rawSpecifier) continue;
      const parsed = parseWorkspaceSpecifier(rawSpecifier.split("?")[0].split("#")[0]);
      if (!parsed) continue;
      const info = workspacePackages.get(parsed.packageName);
      if (!info) continue;
      const resolved = resolveExportPath(info.exports, parsed.exportKey, info.main);
      if (!resolved) continue;
      const resolvedBase = resolved.split("?")[0].split("#")[0];
      for (const ext of disallowedEntrypointExtensions) {
        if (resolvedBase.endsWith(ext)) return true;
      }
    }

    return false;
  }

  for (const file of files) {
    const rawText = await readFile(file, "utf8").catch(() => "");
    const text = stripComments(rawText);
    if (tsImportRe.test(text)) continue;
    if (importsWorkspaceTypeScriptEntrypoint(text)) continue;
    out.push(file);
  }
  return out;
}

/**
 * Load workspace package entrypoints by scanning `package.json` files under the monorepo.
 *
 * This is used for:
 * - skipping node:test suites that depend on TS/TSX entrypoints when TS execution is unavailable, and
 * - resolving missing workspace links in partial `node_modules` environments.
 *
 * @returns {Promise<Map<string, { rootDir: string, exports: any, main: string | null }>>}
 */
/** @type {Promise<Map<string, { rootDir: string, exports: any, main: string | null }>> | null} */
var workspacePackageEntrypointsPromise;
async function loadWorkspacePackageEntrypoints() {
  if (workspacePackageEntrypointsPromise) return workspacePackageEntrypointsPromise;
  workspacePackageEntrypointsPromise = (async () => {
    /** @type {string[]} */
    const packageJsonFiles = [];
    for (const dir of ["packages", "apps", "services", "shared"]) {
      const full = path.join(repoRoot, dir);
      try {
        const stats = await stat(full);
        if (!stats.isDirectory()) continue;
      } catch {
        continue;
      }
      await collectPackageJsonFiles(full, packageJsonFiles);
    }

    /** @type {Map<string, { rootDir: string, exports: any, main: string | null }>} */
    const out = new Map();
    for (const file of packageJsonFiles) {
      const raw = await readFile(file, "utf8").catch(() => null);
      if (!raw) continue;
      let parsed;
      try {
        parsed = JSON.parse(raw);
      } catch {
        continue;
      }
      if (!parsed || typeof parsed !== "object") continue;
      if (typeof parsed.name !== "string") continue;
      if (out.has(parsed.name)) continue;
      out.set(parsed.name, {
        rootDir: path.dirname(file),
        exports: parsed.exports ?? null,
        main: typeof parsed.main === "string" ? parsed.main : null,
      });
    }

    return out;
  })();
  return workspacePackageEntrypointsPromise;
}

/**
 * @param {string} dir
 * @param {string[]} out
 */
async function collectPackageJsonFiles(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    // Keep this list in sync with `scripts/run-node-tests.mjs` at the repo root.
    if (
      entry.name === ".git" ||
      entry.name === "node_modules" ||
      entry.name === "dist" ||
      entry.name === "coverage" ||
      entry.name === "target" ||
      entry.name === "build" ||
      entry.name === ".turbo" ||
      entry.name === ".pnpm-store" ||
      entry.name === ".cache" ||
      entry.name === ".vite" ||
      entry.name === "playwright-report" ||
      entry.name === "test-results" ||
      entry.name === "security-report" ||
      entry.name.startsWith(".tmp")
    ) {
      continue;
    }

    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectPackageJsonFiles(fullPath, out);
      continue;
    }

    if (!entry.isFile()) continue;
    if (entry.name !== "package.json") continue;
    out.push(fullPath);
  }
}

function resolveTypeScriptLoaderArgs() {
  try {
    require.resolve("typescript", { paths: [desktopRoot, repoRoot] });
  } catch {
    return [];
  }

  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-loader.mjs")).href;
  return resolveNodeLoaderArgs(loaderUrl);
}

function getBuiltInTypeScriptSupport() {
  // Prefer explicit flag support when available (older Node versions require it).
  const flagProbe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  if (flagProbe.status === 0) {
    return { enabled: true, args: ["--experimental-strip-types"] };
  }

  const tmpFile = path.join(os.tmpdir(), `formula-strip-types-probe.${process.pid}.${Date.now()}.ts`);
  try {
    writeFileSync(
      tmpFile,
      [
        "export const x: number = 1;",
        "if (x !== 1) throw new Error('strip-types probe failed');",
        "",
      ].join("\n"),
      "utf8",
    );
    const fileUrl = pathToFileURL(tmpFile).href;
    const nativeProbe = spawnSync(process.execPath, ["--input-type=module", "-e", `import ${JSON.stringify(fileUrl)};`], {
      stdio: "ignore",
    });
    if (nativeProbe.status === 0) {
      return { enabled: true, args: [] };
    }
  } catch {
    // ignore
  } finally {
    rmSync(tmpFile, { force: true });
  }

  return { enabled: false, args: [] };
}

/**
 * Resolve Node CLI flags to install an ESM loader.
 *
 * Prefer the newer `register()` API when available (via `--import`), since Node is
 * actively deprecating/removing the older `--experimental-loader` mechanism.
 *
 * @param {string} loaderUrl absolute file:// URL
 * @returns {string[]}
 */
function resolveNodeLoaderArgs(loaderUrl) {
  const allowedFlags =
    process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
      ? process.allowedNodeEnvironmentFlags
      : new Set();

  // `--import` exists before `module.register()` did, so gate on both.
  let supportsRegister = false;
  try {
    supportsRegister = typeof require("node:module")?.register === "function";
  } catch {
    supportsRegister = false;
  }

  if (supportsRegister && allowedFlags.has("--import")) {
    const registerScript = `import { register } from \"node:module\"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) return ["--loader", loaderUrl];
  if (allowedFlags.has("--experimental-loader")) return [`--experimental-loader=${loaderUrl}`];
  return [];
}

async function hasNodeModules() {
  // Prefer the repo-root `node_modules` (pnpm workspace root), but also allow running
  // this script in unusual layouts where deps are installed under `apps/desktop`.
  const candidates = [path.join(repoRoot, "node_modules"), path.join(desktopRoot, "node_modules")];
  let found = false;
  for (const candidate of candidates) {
    try {
      const stats = await stat(candidate);
      if (stats.isDirectory()) {
        found = true;
        break;
      }
    } catch {
      // continue
    }
  }
  if (!found) return false;

  // Some environments may create a minimal `node_modules/` containing only workspace links.
  // Probe for a representative third-party dependency so we can skip tests that require
  // external packages when deps aren't installed.
  try {
    require.resolve("esbuild", { paths: [desktopRoot, repoRoot] });
    return true;
  } catch {
    return false;
  }
}

/**
 * Filter out node:test suites that depend on third-party dependencies when `node_modules`
 * is not installed.
 *
 * @param {string[]} files
 * @param {{ canStripTypes: boolean, canExecuteTsx: boolean }} opts
 */
async function filterExternalDependencyTests(files, opts) {
  /** @type {Map<string, boolean>} */
  const dependencyCache = new Map();
  /**
   * Map an arbitrary directory to the nearest enclosing package.json (if any).
   * The cache stores the *resolved* nearest package for that directory (not just
   * whether that directory itself contains a package.json), so nested dirs inside a
   * workspace package can still resolve package self-references.
   *
   * @type {Map<string, { rootDir: string, name: string | null, exports: any, main: string | null } | null>}
   */
  const packageInfoCache = new Map();
  /** @type {Set<string>} */
  const visiting = new Set();
  const builtins = new Set(builtinModules);
  /** @type {Map<string, { rootDir: string, exports: any, main: string | null }> | null} */
  let workspacePackages = null;

  // Treat `import type ... from "..."` / `export type ... from "..."` as *type-only* when
  // deciding whether a test can run without external deps. These statements are erased by
  // TypeScript (and Node's built-in "strip types" TS execution) and do not create runtime
  // module dependencies.
  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\s*\(\s*["']([^"']+)["']\s*(?:\)|,)/g;
  const dynamicImportTemplateRe = /\bimport\s*\(\s*`((?:\\.|[^`$])*)/g;
  const requireCallRe = /\brequire\(\s*["']([^"']+)["']\s*\)/g;
  const requireResolveRe = /\brequire\.resolve\(\s*["']([^"']+)["']\s*\)/g;
  // Some modules are loaded indirectly via Worker thread entrypoints:
  //   const WORKER_URL = new URL("./sandbox-worker.node.js", import.meta.url)
  // These should be treated as dependencies when deciding which node:test files can run
  // without `node_modules` installed.
  const importMetaUrlRe = /\bnew\s+URL\(\s*["']([^"']+)["']\s*,\s*import\.meta\.url\s*\)/g;

  const candidateExtensions = opts.canExecuteTsx
    ? [".js", ".ts", ".mjs", ".cjs", ".jsx", ".tsx", ".json"]
    : [".js", ".ts", ".mjs", ".cjs", ".jsx", ".json"];

  function parseWorkspaceSpecifier(specifier) {
    if (!specifier.startsWith("@formula/")) return null;
    const parts = specifier.split("/");
    if (parts.length < 2) return null;
    const packageName = `${parts[0]}/${parts[1]}`;
    const subpath = parts.slice(2).join("/");
    const exportKey = subpath ? `./${subpath}` : ".";
    return { packageName, exportKey };
  }

  /**
   * @param {any} exportsMap
   * @param {string} exportKey
   * @param {string | null} main
   */
  function resolveExportPath(exportsMap, exportKey, main) {
    /**
     * @param {any} entry
     * @returns {string | null}
     */
    function pickExportPath(entry) {
      if (typeof entry === "string") return entry;
      if (Array.isArray(entry)) {
        for (const item of entry) {
          const picked = pickExportPath(item);
          if (picked) return picked;
        }
        return null;
      }
      if (entry && typeof entry === "object") {
        // Prefer Node's default ESM condition order.
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.require === "string") return entry.require;

        // Fall back to searching nested condition objects.
        for (const value of Object.values(entry)) {
          const picked = pickExportPath(value);
          if (picked) return picked;
        }
      }
      return null;
    }

    let target = null;
    if (exportsMap) {
      if (typeof exportsMap === "string") {
        if (exportKey === ".") target = exportsMap;
      } else if (exportsMap && typeof exportsMap === "object") {
        const keys = Object.keys(exportsMap);
        const looksLikeSubpathMap = keys.some((k) => k === "." || k.startsWith("./"));
        if (looksLikeSubpathMap) {
          target = pickExportPath(exportsMap?.[exportKey]);
        } else if (exportKey === ".") {
          target = pickExportPath(exportsMap);
        }
      }
    }

    // Packages without an `exports` map (like `@formula/marketplace-shared`) still allow deep
    // imports via `@scope/pkg/<path>`. When analyzing dependencies (including in environments
    // without `node_modules`), treat those subpaths as resolvable workspace files.
    if (!exportsMap && exportKey !== ".") return exportKey;

    if (!target && exportKey === "." && typeof main === "string") target = main;
    return target;
  }

  async function getWorkspacePackages() {
    if (workspacePackages) return workspacePackages;
    workspacePackages = await loadWorkspacePackageEntrypoints();
    return workspacePackages;
  }

  function isBuiltin(specifier) {
    if (specifier.startsWith("node:")) return true;
    return builtins.has(specifier);
  }

  async function nearestPackageInfo(startDir) {
    let dir = startDir;
    /** @type {string[]} */
    const visited = [];
    while (true) {
      const cached = packageInfoCache.get(dir);
      if (cached !== undefined) {
        for (const entry of visited) packageInfoCache.set(entry, cached);
        return cached;
      }

      visited.push(dir);

      const candidate = path.join(dir, "package.json");
      try {
        const raw = await readFile(candidate, "utf8");
        const parsed = JSON.parse(raw);
        const info = {
          rootDir: dir,
          name: typeof parsed?.name === "string" ? parsed.name : null,
          exports: parsed?.exports ?? null,
          main: typeof parsed?.main === "string" ? parsed.main : null,
        };
        for (const entry of visited) packageInfoCache.set(entry, info);
        return info;
      } catch {
        // ignore; walk upward
      }

      if (dir === repoRoot) {
        for (const entry of visited) packageInfoCache.set(entry, null);
        return null;
      }
      const parent = path.dirname(dir);
      if (parent === dir) {
        for (const entry of visited) packageInfoCache.set(entry, null);
        return null;
      }
      dir = parent;
    }
  }

  /**
   * Resolve Node.js package self-references (e.g. `@formula/scripting/node`) to a
   * concrete file path so we can analyze its transitive dependencies.
   *
   * Node supports these self-references even without `node_modules/` when the
   * importing file is inside the package boundary.
   *
   * @param {string} specifier
   * @param {string} importingFile
   */
  async function resolveSelfReference(specifier, importingFile) {
    if (isBuiltin(specifier)) return null;

    const pkgInfo = await nearestPackageInfo(path.dirname(importingFile));
    if (!pkgInfo?.name || !pkgInfo.rootDir) return null;
    if (specifier !== pkgInfo.name && !specifier.startsWith(`${pkgInfo.name}/`)) return null;

    const exportKey = specifier === pkgInfo.name ? "." : `./${specifier.slice(pkgInfo.name.length + 1)}`;

    /**
     * @param {any} entry
     * @returns {string | null}
     */
    function pickExportPath(entry) {
      if (typeof entry === "string") return entry;
      if (Array.isArray(entry)) {
        for (const item of entry) {
          const picked = pickExportPath(item);
          if (picked) return picked;
        }
        return null;
      }
      if (entry && typeof entry === "object") {
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.browser === "string") return entry.browser;

        // Fall back to searching nested condition objects.
        for (const value of Object.values(entry)) {
          const picked = pickExportPath(value);
          if (picked) return picked;
        }
      }
      return null;
    }

    let target = null;
    const exportsMap = pkgInfo.exports;

    if (exportsMap) {
      if (typeof exportsMap === "string") {
        if (exportKey === ".") target = exportsMap;
      } else if (exportsMap && typeof exportsMap === "object") {
        target = pickExportPath(exportsMap?.[exportKey]);
      }
    }

    if (!target && exportKey === "." && pkgInfo.main) target = pkgInfo.main;
    if (!target) return null;

    const resolved = path.resolve(pkgInfo.rootDir, target);
    try {
      const stats = await stat(resolved);
      if (!stats.isFile()) return null;
      return resolved;
    } catch {
      return null;
    }
  }

  async function resolveRelativeModule(importingFile, specifier) {
    const base = path.resolve(path.dirname(importingFile), specifier.split("?")[0].split("#")[0]);
    const ext = path.extname(base);
    if (ext) {
      try {
        const stats = await stat(base);
        if (stats.isFile()) return base;
      } catch {
        // Bundler-style TS sources often import `./foo.js` while the file on disk is
        // `foo.ts` / `foo.tsx`. When TypeScript execution is enabled (and we install the
        // `.js` -> `.ts` resolver loader), treat those as resolvable dependencies for
        // dependency analysis too.
        if (opts.canStripTypes && (ext === ".js" || ext === ".jsx")) {
          const baseNoExt = base.slice(0, -ext.length);
          /** @type {string[]} */
          const fallbacks = [];
          if (ext === ".jsx") {
            // `.jsx` specifiers may point at `.ts` sources (and are supported by the
            // TypeScript transpile loader via `.tsx` -> `.ts` fallbacks). When using Node's
            // built-in TS execution, `.tsx` is unsupported, but `.ts` still works.
            if (opts.canExecuteTsx) fallbacks.push(".tsx");
            fallbacks.push(".ts");
          } else {
            fallbacks.push(".ts");
            if (opts.canExecuteTsx) fallbacks.push(".tsx");
          }

          for (const fallbackExt of fallbacks) {
            const candidate = `${baseNoExt}${fallbackExt}`;
            try {
              const candidateStats = await stat(candidate);
              if (candidateStats.isFile()) return candidate;
            } catch {
              // continue
            }
          }
        }
        return null;
      }
      return null;
    }

    for (const candidateExt of candidateExtensions) {
      const candidate = `${base}${candidateExt}`;
      try {
        const stats = await stat(candidate);
        if (stats.isFile()) return candidate;
      } catch {
        // continue
      }
    }

    // Directory import: try index files.
    try {
      const stats = await stat(base);
      if (stats.isDirectory()) {
        for (const candidateExt of candidateExtensions) {
          const candidate = path.join(base, `index${candidateExt}`);
          try {
            const idxStats = await stat(candidate);
            if (idxStats.isFile()) return candidate;
          } catch {
            // continue
          }
        }
      }
    } catch {
      // ignore
    }

    return null;
  }

  async function fileHasExternalDependencies(file) {
    const cached = dependencyCache.get(file);
    if (cached !== undefined) return cached;
    if (visiting.has(file)) return false;

    visiting.add(file);
    let hasExternal = false;
    const rawText = await readFile(file, "utf8").catch(() => "");
    const text = stripComments(rawText);

    /** @type {string[]} */
    const specifiers = [];
    for (const match of text.matchAll(importFromRe)) {
      const typeOnly = Boolean(match[1]);
      const specifier = match[2];
      if (!specifier) continue;
      if (typeOnly) continue;
      specifiers.push(specifier);
    }
    for (const match of text.matchAll(sideEffectImportRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(dynamicImportRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(dynamicImportTemplateRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(requireCallRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(requireResolveRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(importMetaUrlRe)) {
      const specifier = match[1];
      if (!specifier) continue;
      if (!specifier.startsWith(".") && !specifier.startsWith("/") && !/^[a-zA-Z]+:/.test(specifier)) {
        // `new URL("assets/foo", import.meta.url)` is still relative to the current module.
        specifiers.push(`./${specifier}`);
      } else {
        specifiers.push(specifier);
      }
    }

    for (const specifier of specifiers) {
      if (!specifier) continue;
      if (specifier.startsWith(".")) {
        const resolved = await resolveRelativeModule(file, specifier);
        if (!resolved) continue;
        if (await fileHasExternalDependencies(resolved)) {
          hasExternal = true;
          break;
        }
        continue;
      }

      // Ignore absolute paths and URL-style imports when running node tests.
      if (specifier.startsWith("/") || /^[a-zA-Z]+:/.test(specifier)) continue;

      if (isBuiltin(specifier)) continue;

      const selfResolved = await resolveSelfReference(specifier, file);
      if (selfResolved) {
        if (await fileHasExternalDependencies(selfResolved)) {
          hasExternal = true;
          break;
        }
        continue;
      }

      // Workspace packages can still be imported directly from source even when `node_modules`
      // is missing. Treat those as internal dependencies and recurse.
      if (specifier.startsWith("@formula/")) {
        const parsed = parseWorkspaceSpecifier(specifier.split("?")[0].split("#")[0]);
        if (parsed) {
          const pkgs = await getWorkspacePackages();
          const info = pkgs.get(parsed.packageName);
          if (info) {
            const target = resolveExportPath(info.exports, parsed.exportKey, info.main);
            if (target) {
              const cleanedTarget =
                target.startsWith("./") || target.startsWith("../") || target.startsWith("/")
                  ? target
                  : `./${target}`;
              const basePath = path.resolve(info.rootDir, cleanedTarget.split("?")[0].split("#")[0]);
              let resolved = null;

              if (path.extname(basePath)) {
                try {
                  const stats = await stat(basePath);
                  if (stats.isFile()) resolved = basePath;
                } catch {
                  resolved = null;
                }
              } else {
                for (const ext of candidateExtensions) {
                  const candidate = `${basePath}${ext}`;
                  try {
                    const stats = await stat(candidate);
                    if (stats.isFile()) {
                      resolved = candidate;
                      break;
                    }
                  } catch {
                    // continue
                  }
                }
                if (!resolved) {
                  try {
                    const stats = await stat(basePath);
                    if (stats.isDirectory()) {
                      for (const ext of candidateExtensions) {
                        const candidate = path.join(basePath, `index${ext}`);
                        try {
                          const idxStats = await stat(candidate);
                          if (idxStats.isFile()) {
                            resolved = candidate;
                            break;
                          }
                        } catch {
                          // continue
                        }
                      }
                    }
                  } catch {
                    // ignore
                  }
                }
              }

              if (resolved) {
                if (!opts.canStripTypes && /\.(ts|tsx)$/.test(resolved)) {
                  hasExternal = true;
                  break;
                }
                if (await fileHasExternalDependencies(resolved)) {
                  hasExternal = true;
                  break;
                }
                continue;
              }
            }
          }
        }
      }

      // Any other bare specifier requires external packages.
      hasExternal = true;
      break;
    }

    visiting.delete(file);
    dependencyCache.set(file, hasExternal);
    return hasExternal;
  }

  /** @type {string[]} */
  const out = [];
  for (const file of files) {
    if (await fileHasExternalDependencies(file)) continue;
    out.push(file);
  }
  return out;
}

/**
 * Filter out node:test files that depend on workspace packages that can't be resolved.
 *
 * Some environments (including agent sandboxes) may have third-party dependencies
 * installed but only a subset of workspace package links. In that case, running the
 * full node:test suite would fail fast with `ERR_MODULE_NOT_FOUND` for the missing
 * workspace packages.
 *
 * This function first checks `require.resolve()` (so normal `node_modules` installs
 * keep working), then falls back to a best-effort resolution via local workspace
 * `package.json` `exports`/`main` entries. We skip tests only when a dependency still
 * can't be resolved.
 *
 * @param {string[]} files
 * @param {{ canStripTypes: boolean, canExecuteTsx: boolean }} opts
 */
async function filterMissingWorkspaceDependencyTests(files, opts) {
  /** @type {Map<string, boolean>} */
  const missingCache = new Map();
  /** @type {Set<string>} */
  const visiting = new Set();
  const builtins = new Set(builtinModules);
  /** @type {Map<string, { rootDir: string, exports: any, main: string | null }> | null} */
  let workspacePackages = null;
  /**
   * Map an arbitrary directory to the nearest enclosing package.json (if any).
   * The cache stores the *resolved* nearest package for that directory (not just
   * whether that directory itself contains a package.json), so nested dirs inside a
   * workspace package can still resolve package self-references.
   *
   * @type {Map<string, { rootDir: string, name: string | null, exports: any, main: string | null } | null>}
   */
  const packageInfoCache = new Map();

  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\s*\(\s*["']([^"']+)["']\s*(?:\)|,)/g;
  const dynamicImportTemplateRe = /\bimport\s*\(\s*`((?:\\.|[^`$])*)/g;
  const requireRe = /\brequire\(\s*["']([^"']+)["']\s*\)/g;
  const candidateExtensions = opts.canExecuteTsx
    ? [".js", ".ts", ".mjs", ".cjs", ".jsx", ".tsx", ".json"]
    : [".js", ".ts", ".mjs", ".cjs", ".jsx", ".json"];

  function parseWorkspaceSpecifier(specifier) {
    if (!specifier.startsWith("@formula/")) return null;
    const parts = specifier.split("/");
    if (parts.length < 2) return null;
    const packageName = `${parts[0]}/${parts[1]}`;
    const subpath = parts.slice(2).join("/");
    const exportKey = subpath ? `./${subpath}` : ".";
    return { packageName, exportKey };
  }

  /**
   * @param {any} exportsMap
   * @param {string} exportKey
   * @param {string | null} main
   */
  function resolveExportPath(exportsMap, exportKey, main) {
    /**
     * @param {any} entry
     * @returns {string | null}
     */
    function pickExportPath(entry) {
      if (typeof entry === "string") return entry;
      if (Array.isArray(entry)) {
        for (const item of entry) {
          const picked = pickExportPath(item);
          if (picked) return picked;
        }
        return null;
      }
      if (entry && typeof entry === "object") {
        // Prefer Node's default ESM condition order.
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.require === "string") return entry.require;

        // Fall back to searching nested condition objects.
        for (const value of Object.values(entry)) {
          const picked = pickExportPath(value);
          if (picked) return picked;
        }
      }
      return null;
    }

    let target = null;
    if (exportsMap) {
      if (typeof exportsMap === "string") {
        if (exportKey === ".") target = exportsMap;
      } else if (exportsMap && typeof exportsMap === "object") {
        const keys = Object.keys(exportsMap);
        const looksLikeSubpathMap = keys.some((k) => k === "." || k.startsWith("./"));
        if (looksLikeSubpathMap) {
          target = pickExportPath(exportsMap?.[exportKey]);
        } else if (exportKey === ".") {
          target = pickExportPath(exportsMap);
        }
      }
    }

    // Packages without an `exports` map (like `@formula/marketplace-shared`) still allow deep
    // imports via `@scope/pkg/<path>`. When a workspace link is missing from `node_modules`,
    // fall back to resolving those subpaths directly from the package root directory.
    if (!exportsMap && exportKey !== ".") return exportKey;

    if (!target && exportKey === "." && typeof main === "string") target = main;
    return target;
  }

  async function getWorkspacePackages() {
    if (workspacePackages) return workspacePackages;
    workspacePackages = await loadWorkspacePackageEntrypoints();
    return workspacePackages;
  }

  function isBuiltin(specifier) {
    if (specifier.startsWith("node:")) return true;
    return builtins.has(specifier);
  }

  async function nearestPackageInfo(startDir) {
    let dir = startDir;
    /** @type {string[]} */
    const visited = [];
    while (true) {
      const cached = packageInfoCache.get(dir);
      if (cached !== undefined) {
        for (const entry of visited) packageInfoCache.set(entry, cached);
        return cached;
      }

      visited.push(dir);

      const candidate = path.join(dir, "package.json");
      try {
        const raw = await readFile(candidate, "utf8");
        const parsed = JSON.parse(raw);
        const info = {
          rootDir: dir,
          name: typeof parsed?.name === "string" ? parsed.name : null,
          exports: parsed?.exports ?? null,
          main: typeof parsed?.main === "string" ? parsed.main : null,
        };
        for (const entry of visited) packageInfoCache.set(entry, info);
        return info;
      } catch {
        // ignore; walk upward
      }

      if (dir === repoRoot) {
        for (const entry of visited) packageInfoCache.set(entry, null);
        return null;
      }
      const parent = path.dirname(dir);
      if (parent === dir) {
        for (const entry of visited) packageInfoCache.set(entry, null);
        return null;
      }
      dir = parent;
    }
  }

  /**
   * @param {string} importingFile
   * @param {string} specifier
   */
  async function resolveRelativeModule(importingFile, specifier) {
    const base = path.resolve(path.dirname(importingFile), specifier.split("?")[0].split("#")[0]);
    const ext = path.extname(base);
    if (ext) {
      try {
        const stats = await stat(base);
        if (stats.isFile()) return base;
      } catch {
        // See `scripts/resolve-ts-imports-loader.mjs`: bundler-style `.js` specifiers
        // often point at `.ts`/`.tsx` sources.
        if (opts.canStripTypes && (ext === ".js" || ext === ".jsx")) {
          const baseNoExt = base.slice(0, -ext.length);
          /** @type {string[]} */
          const fallbacks = [];
          if (ext === ".jsx") {
            if (opts.canExecuteTsx) fallbacks.push(".tsx");
            fallbacks.push(".ts");
          } else {
            fallbacks.push(".ts");
            if (opts.canExecuteTsx) fallbacks.push(".tsx");
          }
          for (const fallbackExt of fallbacks) {
            const candidate = `${baseNoExt}${fallbackExt}`;
            try {
              const candidateStats = await stat(candidate);
              if (candidateStats.isFile()) return candidate;
            } catch {
              // continue
            }
          }
        }
        return null;
      }
      return null;
    }

    for (const candidateExt of candidateExtensions) {
      const candidate = `${base}${candidateExt}`;
      try {
        const stats = await stat(candidate);
        if (stats.isFile()) return candidate;
      } catch {
        // continue
      }
    }

    // Directory import: try index files.
    try {
      const stats = await stat(base);
      if (stats.isDirectory()) {
        for (const candidateExt of candidateExtensions) {
          const candidate = path.join(base, `index${candidateExt}`);
          try {
            const idxStats = await stat(candidate);
            if (idxStats.isFile()) return candidate;
          } catch {
            // continue
          }
        }
      }
    } catch {
      // ignore
    }

    return null;
  }

  /**
   * @param {string} specifier
   * @param {string} importingFile
   * @returns {string | null}
   */
  async function resolveWorkspaceSpecifier(specifier, importingFile) {
    // `import.meta.resolve()` currently resolves relative to the module it's invoked from
    // (this runner), so it's not suitable for checking whether a workspace package is
    // installed relative to an arbitrary test file directory. Use `require.resolve()`
    // with explicit `paths` instead so pnpm's per-package `node_modules/` layouts work.
    try {
      return require.resolve(specifier, { paths: [path.dirname(importingFile), repoRoot] });
    } catch {
      // Fall back to a best-effort local workspace resolution. This keeps node:test resilient
      // in environments with stale/partial `node_modules` that are missing some workspace links.
      const parsed = parseWorkspaceSpecifier(specifier.split("?")[0].split("#")[0]);
      if (!parsed) return null;
      const pkgs = await getWorkspacePackages();
      const info = pkgs.get(parsed.packageName);
      if (!info) return null;

      const target = resolveExportPath(info.exports, parsed.exportKey, info.main);
      if (!target) return null;

      const cleanedTarget =
        target.startsWith("./") || target.startsWith("../") || target.startsWith("/") ? target : `./${target}`;
      const basePath = path.resolve(info.rootDir, cleanedTarget.split("?")[0].split("#")[0]);

      if (path.extname(basePath)) {
        try {
          const stats = await stat(basePath);
          if (stats.isFile()) return basePath;
        } catch {
          return null;
        }
        return null;
      }

      for (const ext of candidateExtensions) {
        const candidate = `${basePath}${ext}`;
        try {
          const stats = await stat(candidate);
          if (stats.isFile()) return candidate;
        } catch {
          // continue
        }
      }

      // Directory export: try index files.
      try {
        const stats = await stat(basePath);
        if (stats.isDirectory()) {
          for (const ext of candidateExtensions) {
            const candidate = path.join(basePath, `index${ext}`);
            try {
              const idxStats = await stat(candidate);
              if (idxStats.isFile()) return candidate;
            } catch {
              // continue
            }
          }
        }
      } catch {
        // ignore
      }

      return null;
    }
  }

  /**
   * Resolve Node.js package self-references (e.g. `@formula/python-runtime/native`) to a
   * concrete file path so we can analyze its transitive dependencies.
   *
   * Node supports these self-references even without `node_modules/` when the
   * importing file is inside the package boundary.
   *
   * @param {string} specifier
   * @param {string} importingFile
   */
  async function resolveSelfReference(specifier, importingFile) {
    if (isBuiltin(specifier)) return null;

    const pkgInfo = await nearestPackageInfo(path.dirname(importingFile));
    if (!pkgInfo?.name || !pkgInfo.rootDir) return null;
    if (specifier !== pkgInfo.name && !specifier.startsWith(`${pkgInfo.name}/`)) return null;

    const exportKey = specifier === pkgInfo.name ? "." : `./${specifier.slice(pkgInfo.name.length + 1)}`;

    /**
     * @param {any} entry
     * @returns {string | null}
     */
    function pickExportPath(entry) {
      if (typeof entry === "string") return entry;
      if (Array.isArray(entry)) {
        for (const item of entry) {
          const picked = pickExportPath(item);
          if (picked) return picked;
        }
        return null;
      }
      if (entry && typeof entry === "object") {
        // Node's default export conditions include `node` + `import` with `default` fallback.
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.require === "string") return entry.require;

        // Fall back to searching nested condition objects.
        for (const value of Object.values(entry)) {
          const picked = pickExportPath(value);
          if (picked) return picked;
        }
      }
      return null;
    }

    let target = null;
    const exportsMap = pkgInfo.exports;

    if (exportsMap) {
      if (typeof exportsMap === "string") {
        if (exportKey === ".") target = exportsMap;
      } else if (exportsMap && typeof exportsMap === "object") {
        target = pickExportPath(exportsMap?.[exportKey]);
      }
    }

    if (!target && exportKey === "." && pkgInfo.main) target = pkgInfo.main;
    if (!target) return null;

    const cleanedTarget =
      target.startsWith("./") || target.startsWith("../") || target.startsWith("/") ? target : `./${target}`;
    const basePath = path.resolve(pkgInfo.rootDir, cleanedTarget.split("?")[0].split("#")[0]);
    if (path.extname(basePath)) {
      try {
        const stats = await stat(basePath);
        if (stats.isFile()) return basePath;
      } catch {
        return null;
      }
      return null;
    }

    for (const ext of candidateExtensions) {
      const candidate = `${basePath}${ext}`;
      try {
        const stats = await stat(candidate);
        if (stats.isFile()) return candidate;
      } catch {
        // continue
      }
    }

    // Directory export: try index files.
    try {
      const stats = await stat(basePath);
      if (stats.isDirectory()) {
        for (const ext of candidateExtensions) {
          const candidate = path.join(basePath, `index${ext}`);
          try {
            const idxStats = await stat(candidate);
            if (idxStats.isFile()) return candidate;
          } catch {
            // continue
          }
        }
      }
    } catch {
      // ignore
    }

    return null;
  }

  /**
   * @param {string} file
   * @returns {Promise<boolean>}
   */
  async function fileHasMissingWorkspaceDeps(file) {
    const cached = missingCache.get(file);
    if (cached !== undefined) return cached;
    if (visiting.has(file)) return false;
    visiting.add(file);

    let missing = false;
    const rawText = await readFile(file, "utf8").catch(() => "");
    const text = stripComments(rawText);

    /** @type {Array<{ specifier: string, typeOnly: boolean }>} */
    const imports = [];
    for (const match of text.matchAll(importFromRe)) {
      const typeOnly = Boolean(match[1]);
      const specifier = match[2];
      if (specifier) imports.push({ specifier, typeOnly });
    }
    for (const match of text.matchAll(sideEffectImportRe)) {
      const specifier = match[1];
      if (specifier) imports.push({ specifier, typeOnly: false });
    }
    for (const match of text.matchAll(dynamicImportRe)) {
      const specifier = match[1];
      if (specifier) imports.push({ specifier, typeOnly: false });
    }
    for (const match of text.matchAll(dynamicImportTemplateRe)) {
      const specifier = match[1];
      if (specifier) imports.push({ specifier, typeOnly: false });
    }
    for (const match of text.matchAll(requireRe)) {
      const specifier = match[1];
      if (specifier) imports.push({ specifier, typeOnly: false });
    }

    for (const { specifier: raw, typeOnly } of imports) {
      const specifier = raw.split("?")[0].split("#")[0];
      if (!specifier) continue;

      // Type-only imports are erased by TypeScript stripping and should not be
      // treated as runtime dependencies.
      if (typeOnly && opts.canStripTypes) continue;

      if (specifier.startsWith(".")) {
        const resolved = await resolveRelativeModule(file, specifier);
        if (!resolved) continue;
        if (await fileHasMissingWorkspaceDeps(resolved)) {
          missing = true;
          break;
        }
        continue;
      }

      // Ignore absolute paths and URL-style imports when running node tests.
      if (specifier.startsWith("/") || /^[a-zA-Z]+:/.test(specifier)) continue;
      if (isBuiltin(specifier)) continue;

      // Only check workspace packages; external deps are handled by the normal
      // `node_modules` installation check above.
      if (!specifier.startsWith("@formula/")) continue;

      const selfResolved = await resolveSelfReference(specifier, file);
      if (selfResolved) {
        if (!opts.canStripTypes && /\.(ts|tsx)$/.test(selfResolved)) {
          missing = true;
          break;
        }

        if (selfResolved.startsWith(repoRoot) && (await fileHasMissingWorkspaceDeps(selfResolved))) {
          missing = true;
          break;
        }

        continue;
      }

      const resolved = await resolveWorkspaceSpecifier(specifier, file);
      if (!resolved) {
        missing = true;
        break;
      }

      if (!opts.canStripTypes && /\.(ts|tsx)$/.test(resolved)) {
        missing = true;
        break;
      }

      if (resolved.startsWith(repoRoot) && (await fileHasMissingWorkspaceDeps(resolved))) {
        missing = true;
        break;
      }
    }

    visiting.delete(file);
    missingCache.set(file, missing);
    return missing;
  }

  /** @type {string[]} */
  const out = [];
  for (const file of files) {
    if (await fileHasMissingWorkspaceDeps(file)) continue;
    out.push(file);
  }
  return out;
}
