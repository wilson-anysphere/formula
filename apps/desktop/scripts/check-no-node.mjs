import { readdir, readFile, stat } from "node:fs/promises";
import { builtinModules } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripComments } from "../test/sourceTextUtils.js";

const desktopRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const srcRoot = path.join(desktopRoot, "src");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".mts", ".cts"]);

// The WebView bundle must remain Node-free, but the repo intentionally contains some
// Node-only modules used by Node integration tests (and later bridged via IPC/Tauri).
// Those modules are allowed to import Node built-ins, but *must not* be imported by
// any WebView/runtime code.
//
// Policy: any Node-only tooling should live outside the renderer source tree (`src/`), e.g.:
// - `apps/desktop/tools/**`
// - `apps/desktop/scripts/**`
//
// This check ensures WebView/runtime code under `src/` does not import from those Node-only
// tooling directories.
const NODE_ONLY_PATH_PREFIXES = ["tools/", "scripts/"];
const NODE_ONLY_FILES = [];

const NODE_BUILTIN_PREFIX = "node:";
const NODE_BUILTIN_SPECIFIERS = new Set();
for (const mod of builtinModules) {
  // `builtinModules` is mostly bare specifiers ("fs", "path", "fs/promises", ...)
  // but Node also includes a small set of `node:`-prefixed entries (e.g. `node:sqlite`).
  //
  // Important: do NOT add the stripped `node:` version (e.g. "sqlite") as a banned specifier,
  // because it may refer to a real npm package and would create false positives.
  NODE_BUILTIN_SPECIFIERS.add(mod);
  if (!mod.startsWith(NODE_BUILTIN_PREFIX)) {
    NODE_BUILTIN_SPECIFIERS.add(`${NODE_BUILTIN_PREFIX}${mod}`);
  }
}

const TEST_FILE_RE =
  /\.(test|spec)\.(ts|tsx|js|jsx|mts|cts|mjs|cjs)$|\.vitest\.(ts|tsx|js|jsx|mts|cts|mjs|cjs)$|\.e2e\.(ts|tsx|js|jsx|mts|cts|mjs|cjs)$/;
const IGNORED_DIRS = new Set([
  "node_modules",
  ".git",
  "dist",
  "build",
  "coverage",
  ".turbo",
  ".vite",
  "playwright-report",
  "test-results",
]);

// Match literal specifiers for common import/require forms, including with inline comments
// (e.g. `import(/* webpackChunkName: "x" */ "./foo")`).
const lineComment = "\\/\\/[^\\n]*(?:\\n|$)";
const blockComment = "\\/\\*[\\s\\S]*?\\*\\/";
const wsOrComment = `(?:\\s+|${blockComment}|${lineComment})+`;
const optWsOrComment = `(?:\\s|${blockComment}|${lineComment})*`;

const importFromRe = new RegExp(`\\b(?:import|export)\\s+(?:type\\s+)?[^"']*?\\sfrom${wsOrComment}["']([^"']+)["']`, "g");
const sideEffectImportRe = new RegExp(`\\bimport${wsOrComment}["']([^"']+)["']`, "g");
const dynamicImportRe = new RegExp(`\\bimport\\(${optWsOrComment}["']([^"']+)["']${optWsOrComment}\\)`, "g");
const requireCallRe = new RegExp(`\\brequire\\(${optWsOrComment}["']([^"']+)["']${optWsOrComment}\\)`, "g");

// Catch `process.versions.node` in a few variants (including optional chaining and TS casts),
// but keep it line-local to reduce false positives.
//
// Supported forms include:
// - process.versions.node
// - process?.versions?.node
// - (process as any)?.versions?.node
const processVersionsNodeRe = /\bprocess\b[^\n]{0,120}(?:\.|\?\.)versions(?:\.|\?\.)node\b/g;

/** @type {Map<string, boolean>} */
const isFileCache = new Map();
/** @type {Map<string, boolean>} */
const isDirCache = new Map();

/**
 * @param {string} p
 * @returns {string}
 */
function toPosixPath(p) {
  return p.split(path.sep).join("/");
}

/**
 * @param {string} relPosix
 */
function isTestFile(relPosix) {
  if (relPosix.startsWith("tests/")) return true;
  if (relPosix.includes("/__tests__/")) return true;
  const base = path.posix.basename(relPosix);
  if (
    base.endsWith(".d.ts") ||
    base.endsWith(".d.tsx") ||
    base.endsWith(".d.mts") ||
    base.endsWith(".d.cts")
  ) {
    return true;
  }
  return TEST_FILE_RE.test(base);
}

/**
 * @param {string} relPosix
 */
function isNodeOnlyFile(relPosix) {
  if (NODE_ONLY_FILES.includes(relPosix)) return true;
  return NODE_ONLY_PATH_PREFIXES.some((prefix) => relPosix.startsWith(prefix));
}

/**
 * @param {string} absPath
 * @returns {Promise<boolean>}
 */
async function isFile(absPath) {
  const cached = isFileCache.get(absPath);
  if (cached !== undefined) return cached;
  try {
    const s = await stat(absPath);
    const ok = s.isFile();
    isFileCache.set(absPath, ok);
    return ok;
  } catch {
    isFileCache.set(absPath, false);
    return false;
  }
}

/**
 * @param {string} absPath
 * @returns {Promise<boolean>}
 */
async function isDir(absPath) {
  const cached = isDirCache.get(absPath);
  if (cached !== undefined) return cached;
  try {
    const s = await stat(absPath);
    const ok = s.isDirectory();
    isDirCache.set(absPath, ok);
    return ok;
  } catch {
    isDirCache.set(absPath, false);
    return false;
  }
}

/**
 * Resolve a relative import specifier to an on-disk file, mirroring the Vite
 * "resolve .js -> .ts/.tsx" behavior used by this workspace.
 *
 * @param {string} importerAbs
 * @param {string} rawSpecifier
 * @returns {Promise<string | null>}
 */
async function resolveRelativeImport(importerAbs, rawSpecifier) {
  const cleaned = rawSpecifier.split("?", 1)[0]?.split("#", 1)[0] ?? "";
  if (!cleaned.startsWith(".")) return null;

  const base = path.resolve(path.dirname(importerAbs), cleaned);
  const ext = path.extname(base);

  // Explicit extension.
  if (ext) {
    if (await isFile(base)) return base;

    // Common pattern in this repo: TS sources import `./foo.js` even when the
    // source file is `foo.ts` / `foo.tsx`.
    if (ext === ".js") {
      const ts = base.slice(0, -3) + ".ts";
      if (await isFile(ts)) return ts;
      const tsx = base.slice(0, -3) + ".tsx";
      if (await isFile(tsx)) return tsx;
      const jsx = base.slice(0, -3) + ".jsx";
      if (await isFile(jsx)) return jsx;
    }

    return null;
  }

  // Extensionless import.
  for (const candidateExt of [".ts", ".tsx", ".js", ".jsx"]) {
    const candidate = `${base}${candidateExt}`;
    if (await isFile(candidate)) return candidate;
  }

  // Directory import.
  if (await isDir(base)) {
    for (const candidateExt of [".ts", ".tsx", ".js", ".jsx"]) {
      const candidate = path.join(base, `index${candidateExt}`);
      if (await isFile(candidate)) return candidate;
    }
  }

  return null;
}

/**
 * @param {string} specifier
 * @returns {string | null}
 */
function classifyNodeBuiltinSpecifier(specifier) {
  const cleaned = specifier.split("?", 1)[0]?.split("#", 1)[0] ?? "";
  if (cleaned.startsWith(NODE_BUILTIN_PREFIX)) {
    return `imports Node built-in via "${NODE_BUILTIN_PREFIX}" scheme: "${cleaned}"`;
  }

  if (NODE_BUILTIN_SPECIFIERS.has(cleaned)) {
    return `imports Node built-in module "${cleaned}"`;
  }

  return null;
}

/**
 * @param {string} text
 * @param {number} index
 */
function lineInfo(text, index) {
  const start = text.lastIndexOf("\n", index - 1) + 1;
  const end = text.indexOf("\n", index);
  const line = text.slice(0, start).split("\n").length;
  const snippet = text.slice(start, end === -1 ? text.length : end).trim();
  return { line, snippet };
}

/**
 * @param {string} text
 * @returns {Array<{ specifier: string, index: number }>}
 */
function collectImportSpecifiers(text) {
  /** @type {Array<{ specifier: string, index: number }>} */
  const out = [];

  for (const match of text.matchAll(importFromRe)) {
    const spec = match[1];
    if (!spec) continue;
    const idx = (match.index ?? 0) + match[0].indexOf(spec);
    out.push({ specifier: spec, index: idx });
  }
  for (const match of text.matchAll(sideEffectImportRe)) {
    const spec = match[1];
    if (!spec) continue;
    const idx = (match.index ?? 0) + match[0].indexOf(spec);
    out.push({ specifier: spec, index: idx });
  }
  for (const match of text.matchAll(dynamicImportRe)) {
    const spec = match[1];
    if (!spec) continue;
    const idx = (match.index ?? 0) + match[0].indexOf(spec);
    out.push({ specifier: spec, index: idx });
  }
  for (const match of text.matchAll(requireCallRe)) {
    const spec = match[1];
    if (!spec) continue;
    const idx = (match.index ?? 0) + match[0].indexOf(spec);
    out.push({ specifier: spec, index: idx });
  }

  return out;
}

/**
 * @param {string} absDir
 * @param {Array<{ abs: string, rel: string }>} out
 * @returns {Promise<void>}
 */
async function collectSourceFiles(absDir, out) {
  const entries = await readdir(absDir, { withFileTypes: true });
  for (const entry of entries) {
    const name = entry.name;
    if (IGNORED_DIRS.has(name)) continue;

    const abs = path.join(absDir, name);
    const rel = toPosixPath(path.relative(desktopRoot, abs));

    if (entry.isDirectory()) {
      if (name === "__tests__") continue;
      await collectSourceFiles(abs, out);
      continue;
    }

    if (!entry.isFile()) continue;

    const ext = path.extname(name);
    if (!SOURCE_EXTS.has(ext)) continue;
    if (isTestFile(rel)) continue;

    out.push({ abs, rel });
  }
}

const files = [];
await collectSourceFiles(srcRoot, files);

/** @type {Array<{ abs: string, rel: string }>} */
const runtimeFiles = [];
/** @type {Array<{ abs: string, rel: string }>} */
const nodeOnlyFiles = [];

for (const f of files) {
  if (isNodeOnlyFile(f.rel)) nodeOnlyFiles.push(f);
  else runtimeFiles.push(f);
}

const nodeOnlyAbsFiles = new Set(nodeOnlyFiles.map((f) => f.abs));
// Normalize Node-only directories (strip trailing slashes) so prefix checks work reliably.
const nodeOnlyAbsDirs = NODE_ONLY_PATH_PREFIXES.map((prefix) => path.resolve(desktopRoot, prefix));
const nodeOnlyImportHints = NODE_ONLY_FILES.map((file) =>
  path.posix.basename(file).replace(/\.[^/.]+$/, ""),
);

/** @type {Array<{ file: string, line: number, message: string, snippet: string }>} */
const violations = [];

for (const file of runtimeFiles) {
  const text = await readFile(file.abs, "utf8");
  const scanText = stripComments(text);

  // Node built-ins (direct).
  for (const ref of collectImportSpecifiers(scanText)) {
    const reason = classifyNodeBuiltinSpecifier(ref.specifier);
    if (!reason) continue;
    const { line, snippet } = lineInfo(scanText, ref.index);
    violations.push({
      file: file.rel,
      line,
      message: reason,
      snippet,
    });
  }

  // Node-only runtime detection.
  for (const match of scanText.matchAll(processVersionsNodeRe)) {
    const idx = match.index ?? 0;
    const { line, snippet } = lineInfo(scanText, idx);
    violations.push({
      file: file.rel,
      line,
      message: "uses Node-only runtime API `process.versions.node`",
      snippet,
    });
  }

  // Importing Node-only desktop modules into runtime code.
  for (const ref of collectImportSpecifiers(scanText)) {
    const specifier = ref.specifier.split("?", 1)[0]?.split("#", 1)[0] ?? "";
    if (!specifier.startsWith(".")) continue;

    const baseAbs = path.resolve(path.dirname(file.abs), specifier);
    const couldTargetNodeOnlyDir = nodeOnlyAbsDirs.some((dir) => baseAbs === dir || baseAbs.startsWith(`${dir}${path.sep}`));
    const couldTargetNodeOnlyFile = nodeOnlyImportHints.some((hint) => hint && specifier.includes(hint));
    // Fast path: only resolve specifiers that could plausibly target a known Node-only module.
    // This is intentionally conservative (false positives are ok; false negatives are not).
    if (!couldTargetNodeOnlyDir && !couldTargetNodeOnlyFile) continue;

    const resolved = await resolveRelativeImport(file.abs, specifier);
    if (!resolved) continue;

    const isNodeOnly =
      nodeOnlyAbsFiles.has(resolved) ||
      nodeOnlyAbsDirs.some((dir) => resolved === dir || resolved.startsWith(`${dir}${path.sep}`));
    if (!isNodeOnly) continue;

    const { line, snippet } = lineInfo(scanText, ref.index);
    violations.push({
      file: file.rel,
      line,
      message: `imports Node-only module "${toPosixPath(path.relative(desktopRoot, resolved))}"`,
      snippet,
    });
  }
}

violations.sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line);

if (violations.length > 0) {
  console.error("check-no-node: Node-only APIs detected in desktop WebView/runtime code.\n");
  console.error("The Tauri/WebView bundle must not depend on Node built-ins (fs/path/worker_threads/node:*)\n");
  console.error(`Found ${violations.length} violation(s):\n`);

  for (const v of violations) {
    console.error(`- ${v.file}:${v.line} ${v.message}`);
    if (v.snippet) console.error(`  ${v.snippet}`);
  }

  console.error("\nIf this code is intended to run in Node-only tooling/tests, move it out of the WebView runtime code.");
  console.error("If you need filesystem access at runtime, use Tauri IPC/Rust commands instead of Node built-ins.\n");
  process.exit(1);
}

console.log(`check-no-node: ok (${runtimeFiles.length} runtime file(s) checked)`);
