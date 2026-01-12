import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const desktopRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const srcRoot = path.join(desktopRoot, "src");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx"]);

// The WebView bundle must remain Node-free, but the repo intentionally contains some
// Node-only modules used by Node integration tests (and later bridged via IPC/Tauri).
// Those modules are allowed to import Node built-ins, but *must not* be imported by
// any WebView/runtime code.
// Note: `src/marketplace/` contains runtime-safe WebView code (e.g. URL/config helpers).
// If we ever need Node-only marketplace helpers again, they must live under
// `src/marketplace/node/` (or be added to `NODE_ONLY_FILES`) so the runtime can still
// safely import the shared bits.
const NODE_ONLY_PATH_PREFIXES = ["src/marketplace/node/", "src/security/"];
const NODE_ONLY_FILES = ["src/extensions/ExtensionHostManager.js"];

const NODE_BUILTIN_PREFIX = "node:";
const NODE_BUILTIN_BARE = ["fs", "path", "worker_threads"];

const TEST_FILE_RE = /\.(test|spec)\.[jt]sx?$|\.vitest\.[jt]sx?$|\.e2e\.[jt]sx?$/;
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

const importFromRe = /\b(?:import|export)\s+(?:type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
const dynamicImportRe = /\bimport\(\s*["']([^"']+)["']\s*\)/g;
const requireCallRe = /\brequire\(\s*["']([^"']+)["']\s*\)/g;

// Catch optional chaining variants too (process?.versions?.node, etc).
const processVersionsNodeRe = /\bprocess(?:\?\.)?\.versions(?:\?\.)?\.node\b/g;

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
  if (base.endsWith(".d.ts") || base.endsWith(".d.tsx")) return true;
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

  for (const bare of NODE_BUILTIN_BARE) {
    if (cleaned === bare || cleaned.startsWith(`${bare}/`)) {
      return `imports Node built-in module "${cleaned}"`;
    }
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
const nodeOnlyAbsDirs = NODE_ONLY_PATH_PREFIXES.map((prefix) => path.join(desktopRoot, prefix));

/** @type {Array<{ file: string, line: number, message: string, snippet: string }>} */
const violations = [];

for (const file of runtimeFiles) {
  const text = await readFile(file.abs, "utf8");

  // Node built-ins (direct).
  for (const ref of collectImportSpecifiers(text)) {
    const reason = classifyNodeBuiltinSpecifier(ref.specifier);
    if (!reason) continue;
    const { line, snippet } = lineInfo(text, ref.index);
    violations.push({
      file: file.rel,
      line,
      message: reason,
      snippet,
    });
  }

  // Node-only runtime detection.
  for (const match of text.matchAll(processVersionsNodeRe)) {
    const idx = match.index ?? 0;
    const { line, snippet } = lineInfo(text, idx);
    violations.push({
      file: file.rel,
      line,
      message: "uses Node-only runtime API `process.versions.node`",
      snippet,
    });
  }

  // Importing Node-only desktop modules into runtime code.
  for (const ref of collectImportSpecifiers(text)) {
    const specifier = ref.specifier.split("?", 1)[0]?.split("#", 1)[0] ?? "";
    if (!specifier.startsWith(".")) continue;

    // Fast path: only resolve specifiers that could plausibly target the known Node-only modules.
    if (
      !specifier.includes("marketplace") &&
      !specifier.includes("ExtensionHostManager") &&
      !specifier.includes("security")
    ) {
      continue;
    }

    const resolved = await resolveRelativeImport(file.abs, specifier);
    if (!resolved) continue;

    const isNodeOnly =
      nodeOnlyAbsFiles.has(resolved) ||
      nodeOnlyAbsDirs.some((dir) => resolved === dir || resolved.startsWith(`${dir}${path.sep}`));
    if (!isNodeOnly) continue;

    const { line, snippet } = lineInfo(text, ref.index);
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
