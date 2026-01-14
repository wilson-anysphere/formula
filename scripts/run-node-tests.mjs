import { spawn, spawnSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import { readdir, readFile, stat } from "node:fs/promises";
import { builtinModules, createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { stripComments } from "../apps/desktop/test/sourceTextUtils.js";

// `import.meta.url` is a file URL; use `fileURLToPath` to avoid platform-specific quirks
// (e.g. Windows drive letter handling) when turning it into a filesystem path.
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let cliArgs = process.argv.slice(2);
// pnpm forwards a literal `--` delimiter into scripts. Strip the first occurrence so
// `pnpm test:node -- foo` behaves the same as `pnpm test:node foo`.
const delimiterIdx = cliArgs.indexOf("--");
if (delimiterIdx >= 0) {
  cliArgs = [...cliArgs.slice(0, delimiterIdx), ...cliArgs.slice(delimiterIdx + 1)];
}

if (cliArgs.includes("--help") || cliArgs.includes("-h")) {
  console.log(
    [
      "Run node:test suites in this repo (with safe defaults for multi-agent hosts).",
      "",
      "Usage:",
      "  pnpm test:node",
      "  pnpm test:node <pattern> [pattern...]",
      "",
      "Notes:",
      "  - Additional args are treated as substring filters over test file paths.",
      "  - pnpm forwards a literal `--` delimiter; this script strips it automatically.",
      "",
      "Examples:",
      "  pnpm test:node collab",
      "  pnpm test:node startSyncServer",
      "  pnpm test:node -- desktop",
      "",
    ].join("\n"),
  );
  process.exit(0);
}

const fileFilters = cliArgs
  .filter((arg) => typeof arg === "string")
  .map((arg) => arg.trim())
  .filter((arg) => arg !== "" && !arg.startsWith("-"));

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collect(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    // Skip node_modules and other generated output.
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
      await collect(fullPath, out);
      continue;
    }

    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".test.js")) continue;
    out.push(fullPath);
  }
}

/** @type {string[]} */
const testFiles = [];
await collect(repoRoot, testFiles);
testFiles.sort();

let filteredTestFiles = testFiles;
if (fileFilters.length > 0) {
  const lowered = fileFilters.map((filter) => filter.toLowerCase());
  filteredTestFiles = testFiles.filter((file) => {
    const rel = path.relative(repoRoot, file).split(path.sep).join("/");
    const haystack = rel.toLowerCase();
    return lowered.some((needle) => haystack.includes(needle));
  });
  console.log(
    `[node:test] Filtering ${testFiles.length} test file(s) by ${fileFilters.length} pattern(s): ${fileFilters.join(", ")}`,
  );
}

const tsLoaderArgs = resolveTypeScriptLoaderArgs();
const builtInTypeScript = getBuiltInTypeScriptSupport();
const canExecuteTypeScript = builtInTypeScript.enabled || tsLoaderArgs.length > 0;
// Node's built-in "strip types" support can execute `.ts` modules, but does not support
// `.tsx` (JSX) without a real transpile loader.
const canExecuteTsx = tsLoaderArgs.length > 0;

let runnableTestFiles = testFiles;
let typeScriptFilteredCount = 0;
let typeScriptTsxFilteredCount = 0;
if (!canExecuteTypeScript) {
  runnableTestFiles = await filterTypeScriptImportTests(filteredTestFiles, ["ts", "tsx"]);
  typeScriptFilteredCount = filteredTestFiles.length - runnableTestFiles.length;
} else if (!canExecuteTsx) {
  runnableTestFiles = await filterTypeScriptImportTests(filteredTestFiles, ["tsx"]);
  typeScriptTsxFilteredCount = filteredTestFiles.length - runnableTestFiles.length;
} else {
  runnableTestFiles = filteredTestFiles;
}

const hasDeps = await hasNodeModules();
let externalDepsFilteredCount = 0;
let missingWorkspaceDepsFilteredCount = 0;
if (!hasDeps) {
  const before = runnableTestFiles.length;
  runnableTestFiles = await filterExternalDependencyTests(runnableTestFiles, {
    canStripTypes: canExecuteTypeScript,
    canExecuteTsx,
  });
  externalDepsFilteredCount = before - runnableTestFiles.length;
} else {
  const before = runnableTestFiles.length;
  runnableTestFiles = await filterMissingWorkspaceDependencyTests(runnableTestFiles, { canStripTypes: canExecuteTypeScript, canExecuteTsx });
  missingWorkspaceDepsFilteredCount = before - runnableTestFiles.length;
}

if (runnableTestFiles.length !== filteredTestFiles.length) {
  const skipped = filteredTestFiles.length - runnableTestFiles.length;
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

if (runnableTestFiles.length === 0) {
  if (fileFilters.length > 0) {
    console.log(`No node:test files matched: ${fileFilters.join(", ")}`);
  } else {
    console.log("No node:test files found.");
  }
  process.exit(0);
}

const baseNodeArgs = ["--no-warnings"];
if (tsLoaderArgs.length > 0) {
  baseNodeArgs.push(...tsLoaderArgs);
} else if (builtInTypeScript.enabled) {
  baseNodeArgs.push(...builtInTypeScript.args);
  // Many TS sources in this repo use `.js` specifiers that point at `.ts` files
  // (bundler-style resolution). Node's default ESM resolver does not support that,
  // so we install a tiny loader that falls back from `./foo.js` -> `./foo.ts`.
  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
  baseNodeArgs.push(...resolveNodeLoaderArgs(loaderUrl));
}

// Node's test runner defaults `--test-concurrency` to the number of available CPU
// cores. On large/shared runners this can massively over-parallelize heavyweight
// integration tests (sync-server, extension sandboxing, python runtime) and lead
// to spurious timeouts. Keep a conservative default, with an escape hatch for
// CI/local tuning.
const envConcurrency = Number.parseInt(
  process.env.FORMULA_NODE_TEST_CONCURRENCY ?? process.env.NODE_TEST_CONCURRENCY ?? "",
  10,
);
let testConcurrency = Number.isFinite(envConcurrency) && envConcurrency > 0 ? envConcurrency : 2;
// Clamp to something reasonable even if the env var is set to a huge value.
// (In practice we don't want to spin up hundreds of node:test workers on a big host.)
const maxTestConcurrency = Math.min(16, os.availableParallelism?.() ?? 4);
testConcurrency = Math.max(1, Math.min(testConcurrency, maxTestConcurrency));

// `.e2e.test.js` suites tend to start background services (sync-server, sandbox
// workers, etc). Even with a low global `--test-concurrency`, running those files
// in parallel with unrelated unit tests can still create enough load to cause
// spurious timeouts. Run them separately with `--test-concurrency=1` to keep the
// suite reliable.
const e2eTestFiles = runnableTestFiles.filter((file) => file.endsWith(".e2e.test.js"));
const nonE2eTestFiles = runnableTestFiles.filter((file) => !file.endsWith(".e2e.test.js"));

// Some unit tests still start sync-server child processes (via `startSyncServer`).
// Running multiple of those alongside other suites can be noisy/flaky under load.
// Detect and serialize them as a separate batch.
/** @type {string[]} */
const syncServerTestFiles = [];
/** @type {string[]} */
const unitTestFiles = [];
const syncServerRe = /\bstartSyncServer\b/;
for (const file of nonE2eTestFiles) {
  const text = await readFile(file, "utf8").catch(() => "");
  if (syncServerRe.test(text)) syncServerTestFiles.push(file);
  else unitTestFiles.push(file);
}

/**
 * @param {string[]} files
 * @param {number} concurrency
 * @returns {Promise<number>} exit code
 */
async function runTestBatch(files, concurrency) {
  if (files.length === 0) return 0;

  const nodeArgs = [...baseNodeArgs, `--test-concurrency=${concurrency}`, "--test", ...files];
  const child = spawn(process.execPath, nodeArgs, { stdio: "inherit" });
  return await new Promise((resolve) => {
    child.on("exit", (code, signal) => {
      if (signal) {
        console.error(`node:test exited with signal ${signal}`);
        resolve(1);
        return;
      }
      resolve(code ?? 1);
    });
  });
}

/**
 * Run files serially, each in their own `node --test` invocation.
 *
 * This is intentionally more expensive than batching (extra process startup), but avoids
 * cross-file interference for heavyweight integration suites that start/stop background
 * services (e.g. sync-server) and have been observed to flake/time out when run in a
 * single shared node:test process.
 *
 * @param {string[]} files
 * @param {number} concurrency
 * @returns {Promise<number>} exit code
 */
async function runTestFilesIndividually(files, concurrency) {
  for (const file of files) {
    const code = await runTestBatch([file], concurrency);
    if (code !== 0) return code;
  }
  return 0;
}

let exitCode = 0;
exitCode = await runTestFilesIndividually(e2eTestFiles, 1);
if (exitCode !== 0) process.exit(exitCode);
exitCode = await runTestFilesIndividually(syncServerTestFiles, 1);
if (exitCode !== 0) process.exit(exitCode);
exitCode = await runTestBatch(unitTestFiles, testConcurrency);
process.exit(exitCode);

function getBuiltInTypeScriptSupport() {
  // Prefer explicit flag support when available (older Node versions require it).
  const flagProbe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  if (flagProbe.status === 0) {
    return { enabled: true, args: ["--experimental-strip-types"] };
  }

  // Newer Node versions may support executing `.ts` files without the flag. Probe that
  // behavior by importing a temporary `.ts` module (ESM), since this runner executes
  // test files under `node --test` in ESM mode.
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
    const registerScript = `import { register } from "node:module"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) {
    return ["--loader", loaderUrl];
  }
  if (allowedFlags.has("--experimental-loader")) {
    return [`--experimental-loader=${loaderUrl}`];
  }
  return [];
}

function resolveTypeScriptLoaderArgs() {
  // When `typescript` is available, prefer a real TS->JS transpile loader over Node's
  // "strip-only" TS support:
  // - strip-only mode rejects TS runtime features like parameter properties and enums
  // - many packages use `.js` specifiers that should resolve to `.ts` sources (TS ESM convention)
  //
  // This keeps `node --test` usable without a separate build step.
  try {
    require.resolve("typescript", { paths: [repoRoot] });
  } catch {
    return [];
  }

  const loaderUrl = new URL("./resolve-ts-loader.mjs", import.meta.url).href;
  return resolveNodeLoaderArgs(loaderUrl);
}

async function hasNodeModules() {
  try {
    const stats = await stat(path.join(repoRoot, "node_modules"));
    if (!stats.isDirectory()) return false;
  } catch {
    return false;
  }

  // Some repo environments (including agent sandboxes) may create a minimal `node_modules/`
  // containing only workspace links without actually installing third-party packages.
  // Probe for a representative dependency so we can skip tests that require external
  // packages when deps aren't installed.
  try {
    require.resolve("esbuild", { paths: [repoRoot] });
    return true;
  } catch {
    return false;
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

  // When TypeScript/TSX execution is unavailable (older Node versions without `--experimental-strip-types`
  // or when running under Node's built-in TS execution which does not support `.tsx`), we still want to
  // skip suites that depend on workspace packages whose entrypoints are authored as `.ts`/`.tsx` (even
  // if the test itself doesn't import a `.ts`/`.tsx` file directly).
  //
  // We treat `extensions` as the set of *disallowed* entrypoint extensions for this environment.
  const workspacePackages = await loadWorkspacePackageEntrypoints();
  const disallowedEntrypointExtensions = new Set(extensions.map((ext) => `.${ext}`));
  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\s*\(\s*["']([^"']+)["']\s*(?:\)|,)/g;
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
    // imports via `@scope/pkg/<path>`. When a workspace link is missing from `node_modules`,
    // fall back to resolving those subpaths directly from the package root directory.
    if (!exportsMap && exportKey !== ".") return exportKey;

    if (!target && exportKey === "." && typeof main === "string") target = main;
    return target;
  }

  /**
   * @param {string} text
   */
  function importsWorkspaceTypeScriptEntrypoint(text) {
    if (!workspacePackages) return false;
    /** @type {string[]} */
    const specifiers = [];
    for (const match of text.matchAll(importFromRe)) specifiers.push(match[2]);
    for (const match of text.matchAll(sideEffectImportRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(dynamicImportRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(dynamicImportTemplateRe)) specifiers.push(match[1]);
    for (const match of text.matchAll(requireCallRe)) specifiers.push(match[1]);

    for (const specifier of specifiers) {
      if (!specifier) continue;
      const parsed = parseWorkspaceSpecifier(specifier.split("?")[0].split("#")[0]);
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
    // Heuristic: skip tests that import TypeScript modules when the runtime can't execute them.
    if (tsImportRe.test(text)) continue;
    if (importsWorkspaceTypeScriptEntrypoint(text)) continue;
    out.push(file);
  }
  return out;
}

/**
 * Load workspace package entrypoints by scanning `package.json` files under `packages/`.
 *
 * This is used only when TypeScript execution is unavailable, so we can skip node:test
 * suites that import workspace packages authored as `.ts`/`.tsx`.
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
    // Skip node_modules and other generated output.
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

  // When dependencies are not installed, `node_modules` may be missing (or contain only a
  // subset of workspace links). In that scenario, `@formula/*` workspace imports can still
  // be resolved directly from the repo source tree via `package.json` exports/main.
  const workspacePackages = await loadWorkspacePackageEntrypoints();

  function isBuiltin(specifier) {
    if (specifier.startsWith("node:")) return true;
    return builtins.has(specifier);
  }

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
    // imports via `@scope/pkg/<path>`.
    if (!exportsMap && exportKey !== ".") return exportKey;

    if (!target && exportKey === "." && typeof main === "string") target = main;
    return target;
  }

  /**
   * Resolve a workspace package import (e.g. `@formula/collab-session` or
   * `@formula/marketplace-shared/extension-package/v2-browser.mjs`) directly from the repo.
   *
   * @param {string} specifier
   * @returns {Promise<string | null>} absolute file path
   */
  async function resolveWorkspaceImport(specifier) {
    const parsed = parseWorkspaceSpecifier(specifier.split("?")[0].split("#")[0]);
    if (!parsed) return null;
    const info = workspacePackages.get(parsed.packageName);
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

    const exportKey =
      specifier === pkgInfo.name ? "." : `./${specifier.slice(pkgInfo.name.length + 1)}`;

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
      if (
        !specifier.startsWith(".") &&
        !specifier.startsWith("/") &&
        !/^[a-zA-Z]+:/.test(specifier)
      ) {
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
        const workspaceResolved = await resolveWorkspaceImport(specifier);
        if (workspaceResolved) {
          if (!opts.canStripTypes && /\.(ts|tsx)$/.test(workspaceResolved)) {
            hasExternal = true;
            break;
          }
          if (await fileHasExternalDependencies(workspaceResolved)) {
            hasExternal = true;
            break;
          }
          continue;
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
  const workspacePackages = await loadWorkspacePackageEntrypoints();
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
  // Detect dynamic imports with string literal specifiers so we can follow transitive
  // dependencies when runtime wiring uses `await import("./foo.js")` (common in ESM).
  const dynamicImportRe = /\bimport\s*\(\s*["']([^"']+)["']\s*(?:\)|,)/g;
  const dynamicImportTemplateRe = /\bimport\s*\(\s*`((?:\\.|[^`$])*)/g;
  // Some Node-oriented packages (or older code) still use `require()` even under
  // `node --test` ESM mode. Treat string-literal requires as dependencies too.
  const requireRe = /\brequire\s*\(\s*["']([^"']+)["']\s*\)/g;
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
        if (typeof entry.node === "string") return entry.node;
        if (typeof entry.import === "string") return entry.import;
        if (typeof entry.default === "string") return entry.default;
        if (typeof entry.require === "string") return entry.require;
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
      const info = workspacePackages.get(parsed.packageName);
      if (!info) return null;

      const target = resolveExportPath(info.exports, parsed.exportKey, info.main);
      if (!target) return null;

      const cleanedTarget =
        target.startsWith("./") || target.startsWith("../") || target.startsWith("/")
          ? target
          : `./${target}`;
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

    const exportKey =
      specifier === pkgInfo.name ? "." : `./${specifier.slice(pkgInfo.name.length + 1)}`;

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
      target.startsWith("./") || target.startsWith("../") || target.startsWith("/")
        ? target
        : `./${target}`;
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
