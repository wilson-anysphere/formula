import fs from "node:fs/promises";
import path from "node:path";
import { builtinModules } from "node:module";
import process from "node:process";
import { fileURLToPath } from "node:url";

/**
 * Guardrail: the Tauri desktop renderer bundle must stay Node-free.
 *
 * The WebView runtime does not provide Node built-ins like `fs`, `path`,
 * `worker_threads`, etc. Accidentally importing those from renderer code can
 * slip past review (especially if the module is not currently imported by the
 * entrypoint) and later break at runtime.
 *
 * This script scans `apps/desktop/src/**` (excluding test files) and fails if it:
 * - imports a Node built-in module (e.g. `node:fs`, `path`, `worker_threads`)
 * - imports code from `apps/desktop/tools/**` or `apps/desktop/scripts/**` (Node-only tooling)
 */

// Resolve repo root relative to this script so callers don't have to `cd` first.
// (`pnpm -w lint` runs from the repo root today, but this makes the guard more robust.)
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const desktopSrcDir = path.join(repoRoot, "apps", "desktop", "src");
const desktopToolsDir = path.join(repoRoot, "apps", "desktop", "tools");
const desktopScriptsDir = path.join(repoRoot, "apps", "desktop", "scripts");

const SOURCE_EXTENSIONS = new Set([".js", ".jsx", ".ts", ".tsx", ".mjs", ".cjs"]);

const BANNED_MODULE_SPECIFIERS = new Set();
for (const mod of builtinModules) {
  BANNED_MODULE_SPECIFIERS.add(mod);
  if (mod.startsWith("node:")) {
    BANNED_MODULE_SPECIFIERS.add(mod.slice("node:".length));
  } else {
    BANNED_MODULE_SPECIFIERS.add(`node:${mod}`);
  }
}

function isDesktopRendererSourceFile(absPath) {
  const ext = path.extname(absPath);
  if (!SOURCE_EXTENSIONS.has(ext)) return false;

  const base = path.basename(absPath);
  if (base.includes(".test.") || base.includes(".spec.") || base.includes(".vitest.") || base.includes(".e2e.")) {
    return false;
  }

  const parts = absPath.split(path.sep);
  if (parts.includes("__tests__")) return false;

  return true;
}

async function* walkFiles(dir) {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const absPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === "__tests__") continue;
      yield* walkFiles(absPath);
      continue;
    }
    if (entry.isFile()) {
      yield absPath;
    }
  }
}

function listImportSpecifiers(sourceText) {
  /** @type {{ specifier: string, index: number, kind: string }[]} */
  const out = [];

  const patterns = [
    { kind: "import-from", re: /\bimport\s+(?:type\s+)?[^'"]*?\s+from\s+['"]([^'"]+)['"]/g },
    { kind: "import-side-effect", re: /\bimport\s+['"]([^'"]+)['"]/g },
    { kind: "export-from", re: /\bexport\s+[^'"]*?\s+from\s+['"]([^'"]+)['"]/g },
    { kind: "dynamic-import", re: /\bimport\s*\(\s*['"]([^'"]+)['"]\s*\)/g },
    { kind: "require", re: /\brequire\s*\(\s*['"]([^'"]+)['"]\s*\)/g },
  ];

  for (const { kind, re } of patterns) {
    for (const match of sourceText.matchAll(re)) {
      out.push({ kind, specifier: match[1], index: match.index ?? 0 });
    }
  }

  return out;
}

function lineAndColumnForIndex(sourceText, index) {
  const before = sourceText.slice(0, index);
  const line = before.split(/\r?\n/).length;
  const lastNewline = before.lastIndexOf("\n");
  const column = index - (lastNewline === -1 ? 0 : lastNewline + 1) + 1;
  return { line, column };
}

function isBannedImport(specifier) {
  if (specifier.startsWith("node:")) return true;
  return BANNED_MODULE_SPECIFIERS.has(specifier);
}

function stripQueryAndHash(specifier) {
  const queryIndex = specifier.indexOf("?");
  const hashIndex = specifier.indexOf("#");
  let end = specifier.length;
  if (queryIndex !== -1) end = Math.min(end, queryIndex);
  if (hashIndex !== -1) end = Math.min(end, hashIndex);
  return specifier.slice(0, end);
}

function isPathWithin(absPath, absDir) {
  const rel = path.relative(absDir, absPath);
  return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel));
}

/** @type {{ file: string, line: number, column: number, kind: string, specifier: string }[]} */
const violations = [];

try {
  await fs.access(desktopSrcDir);
} catch {
  console.error(`Expected desktop source directory at ${path.relative(repoRoot, desktopSrcDir)}, but it does not exist.`);
  process.exitCode = 1;
  process.exit(process.exitCode);
}

for await (const absPath of walkFiles(desktopSrcDir)) {
  if (!isDesktopRendererSourceFile(absPath)) continue;

  const sourceText = await fs.readFile(absPath, "utf8");
  const imports = listImportSpecifiers(sourceText);
  for (const imp of imports) {
    // Keep Node-only tooling out of the WebView renderer import graph. Anything in
    // `apps/desktop/tools` or `apps/desktop/scripts` is assumed to be Node-only (or at least
    // not safe to bundle into the renderer).
    if (imp.specifier.startsWith(".")) {
      const cleaned = stripQueryAndHash(imp.specifier);
      const resolved = path.resolve(path.dirname(absPath), cleaned);
      if (isPathWithin(resolved, desktopToolsDir) || isPathWithin(resolved, desktopScriptsDir)) {
        const { line, column } = lineAndColumnForIndex(sourceText, imp.index);
        violations.push({
          file: path.relative(repoRoot, absPath),
          line,
          column,
          kind: `${imp.kind} (renderer-imports-node-only-tooling)`,
          specifier: imp.specifier,
        });
        continue;
      }
    }

    if (!isBannedImport(imp.specifier)) continue;
    const { line, column } = lineAndColumnForIndex(sourceText, imp.index);
    violations.push({
      file: path.relative(repoRoot, absPath),
      line,
      column,
      kind: imp.kind,
      specifier: imp.specifier,
    });
  }
}

if (violations.length > 0) {
  console.error("Desktop renderer must stay Node-free. Found Node built-in imports in apps/desktop/src (excluding tests):");
  for (const v of violations) {
    console.error(`- ${v.file}:${v.line}:${v.column} ${v.kind} -> "${v.specifier}"`);
  }
  console.error("");
  console.error(
    "Move Node-only modules under apps/desktop/scripts/ or apps/desktop/tools/ and keep them out of the renderer import graph."
  );
  process.exitCode = 1;
}
