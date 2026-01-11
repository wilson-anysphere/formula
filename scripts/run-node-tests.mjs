import { spawn, spawnSync } from "node:child_process";
import { readdir, readFile, stat } from "node:fs/promises";
import { builtinModules } from "node:module";
import path from "node:path";

const repoRoot = path.resolve(new URL(".", import.meta.url).pathname, "..");

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collect(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    // Skip node_modules and other generated output.
    if (entry.name === "node_modules" || entry.name === "dist" || entry.name === "coverage") continue;

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

const hasDeps = await hasNodeModules();
if (!hasDeps) {
  runnableTestFiles = await filterExternalDependencyTests(runnableTestFiles);
}

if (runnableTestFiles.length !== testFiles.length) {
  const skipped = testFiles.length - runnableTestFiles.length;
  if (canStripTypes) {
    console.log(`Skipping ${skipped} node:test file(s) that depend on external packages (node_modules missing).`);
  } else if (hasDeps) {
    console.log(`Skipping ${skipped} node:test file(s) that import .ts modules (TypeScript stripping not available).`);
  } else {
    console.log(
      `Skipping ${skipped} node:test file(s) that import .ts modules or depend on external packages (TypeScript stripping and node_modules missing).`,
    );
  }
}

if (runnableTestFiles.length === 0) {
  console.log("No node:test files found.");
  process.exit(0);
}

const nodeArgs = ["--no-warnings"];
if (canStripTypes) nodeArgs.push("--experimental-strip-types");
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
    return stats.isDirectory();
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
  /** @type {Map<string, string | null>} */
  const packageNameCache = new Map();
  /** @type {Set<string>} */
  const visiting = new Set();
  const builtins = new Set(builtinModules);

  const importFromRe = /\b(?:import|export)\s+(?:type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\(\s*["']([^"']+)["']\s*\)/g;

  const candidateExtensions = [".js", ".ts", ".mjs", ".cjs", ".jsx", ".tsx", ".json"];

  function isBuiltin(specifier) {
    if (specifier.startsWith("node:")) return true;
    return builtins.has(specifier);
  }

  async function nearestPackageName(startDir) {
    let dir = startDir;
    while (true) {
      const cached = packageNameCache.get(dir);
      if (cached !== undefined) return cached;

      const candidate = path.join(dir, "package.json");
      try {
        const raw = await readFile(candidate, "utf8");
        const parsed = JSON.parse(raw);
        const name = typeof parsed?.name === "string" ? parsed.name : null;
        packageNameCache.set(dir, name);
        return name;
      } catch {
        packageNameCache.set(dir, null);
      }

      if (dir === repoRoot) return null;
      const parent = path.dirname(dir);
      if (parent === dir) return null;
      dir = parent;
    }
  }

  async function allowsBareSpecifier(specifier, importingFile) {
    if (isBuiltin(specifier)) return true;
    const pkgName = await nearestPackageName(path.dirname(importingFile));
    if (!pkgName) return false;
    return specifier === pkgName || specifier.startsWith(`${pkgName}/`);
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

      if (!(await allowsBareSpecifier(specifier, file))) {
        hasExternal = true;
        break;
      }
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
