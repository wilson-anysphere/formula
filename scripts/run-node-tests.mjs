import { spawn, spawnSync } from "node:child_process";
import { readdir, readFile, stat } from "node:fs/promises";
import { builtinModules, createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const repoRoot = path.resolve(new URL(".", import.meta.url).pathname, "..");
const require = createRequire(import.meta.url);

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

const canStripTypes = supportsTypeStripping();
let runnableTestFiles = canStripTypes ? testFiles : await filterTypeScriptImportTests(testFiles);
const typeScriptFilteredCount = testFiles.length - runnableTestFiles.length;

const hasDeps = await hasNodeModules();
let externalDepsFilteredCount = 0;
let missingWorkspaceDepsFilteredCount = 0;
if (!hasDeps) {
  const before = runnableTestFiles.length;
  runnableTestFiles = await filterExternalDependencyTests(runnableTestFiles);
  externalDepsFilteredCount = before - runnableTestFiles.length;
} else {
  const before = runnableTestFiles.length;
  runnableTestFiles = await filterMissingWorkspaceDependencyTests(runnableTestFiles, { canStripTypes });
  missingWorkspaceDepsFilteredCount = before - runnableTestFiles.length;
}

if (runnableTestFiles.length !== testFiles.length) {
  const skipped = testFiles.length - runnableTestFiles.length;
  /** @type {string[]} */
  const reasons = [];
  if (typeScriptFilteredCount > 0) {
    reasons.push(`${typeScriptFilteredCount} import .ts modules (TypeScript stripping not available)`);
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
  console.log("No node:test files found.");
  process.exit(0);
}

const nodeArgs = ["--no-warnings"];
if (canStripTypes) nodeArgs.push("--experimental-strip-types");
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
nodeArgs.push(`--test-concurrency=${testConcurrency}`);
nodeArgs.push("--test", ...runnableTestFiles);

const child = spawn(process.execPath, nodeArgs, {
  stdio: "inherit",
});

child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

function supportsTypeStripping() {
  const probe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  return probe.status === 0;
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

async function filterTypeScriptImportTests(files) {
  /** @type {string[]} */
  const out = [];
  const tsImportRe = /from\s+["'][^"']+\.(ts|tsx)["']|import\(\s*["'][^"']+\.(ts|tsx)["']\s*\)/;
  for (const file of files) {
    const text = await readFile(file, "utf8").catch(() => "");
    // Heuristic: skip tests that import .ts modules when the runtime can't strip types.
    if (tsImportRe.test(text)) continue;
    out.push(file);
  }
  return out;
}

async function filterExternalDependencyTests(files) {
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

  const importFromRe = /\b(?:import|export)\s+(?:type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\(\s*["']([^"']+)["']\s*\)/g;
  const requireCallRe = /\brequire\(\s*["']([^"']+)["']\s*\)/g;
  const requireResolveRe = /\brequire\.resolve\(\s*["']([^"']+)["']\s*\)/g;
  // Some modules are loaded indirectly via Worker thread entrypoints:
  //   const WORKER_URL = new URL("./sandbox-worker.node.js", import.meta.url)
  // These should be treated as dependencies when deciding which node:test files can run
  // without `node_modules` installed.
  const importMetaUrlRe = /\bnew\s+URL\(\s*["']([^"']+)["']\s*,\s*import\.meta\.url\s*\)/g;

  const candidateExtensions = [".js", ".ts", ".mjs", ".cjs", ".jsx", ".tsx", ".json"];

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
    const text = await readFile(file, "utf8").catch(() => "");

    /** @type {string[]} */
    const specifiers = [];
    for (const match of text.matchAll(importFromRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(sideEffectImportRe)) {
      specifiers.push(match[1]);
    }
    for (const match of text.matchAll(dynamicImportRe)) {
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
 * Filter out node:test files that import workspace packages that aren't present in
 * the local `node_modules/` tree.
 *
 * Some environments (including agent sandboxes) may have third-party dependencies
 * installed but only a subset of workspace package links. In that case, running
 * the full node:test suite would fail fast with `ERR_MODULE_NOT_FOUND` for the
 * missing workspace packages.
 *
 * We conservatively skip tests that depend on missing `@formula/*` imports.
 *
 * @param {string[]} files
 * @param {{ canStripTypes: boolean }} opts
 */
async function filterMissingWorkspaceDependencyTests(files, opts) {
  /** @type {Map<string, boolean>} */
  const missingCache = new Map();
  /** @type {Set<string>} */
  const visiting = new Set();
  const builtins = new Set(builtinModules);

  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const candidateExtensions = [".js", ".ts", ".mjs", ".cjs", ".jsx", ".tsx", ".json"];

  function isBuiltin(specifier) {
    if (specifier.startsWith("node:")) return true;
    return builtins.has(specifier);
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
  function resolveWorkspaceSpecifier(specifier, importingFile) {
    try {
      const parentUrl = pathToFileURL(importingFile).href;
      if (typeof import.meta.resolve === "function") {
        const resolved = import.meta.resolve(specifier, parentUrl);
        if (resolved && resolved.startsWith("file:")) return fileURLToPath(resolved);
        return null;
      }
    } catch {
      return null;
    }

    try {
      return require.resolve(specifier, { paths: [path.dirname(importingFile), repoRoot] });
    } catch {
      return null;
    }
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
    const text = await readFile(file, "utf8").catch(() => "");

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

      const resolved = resolveWorkspaceSpecifier(specifier, file);
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
