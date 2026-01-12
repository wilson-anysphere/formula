#!/usr/bin/env node
/**
 * Cursor-only AI policy guard.
 *
 * This repository is a Cursor product:
 *   - No OpenAI / Anthropic / Ollama integrations
 *   - No user-supplied API keys
 *   - No local model toggles
 *
 * This script is intended as a fast CI regression guard (not a full linter).
 * It scans source-code directories for a small set of forbidden patterns and
 * exits non-zero when any are found.
 */
import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import os from "node:os";

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const DEFAULT_REPO_ROOT = path.resolve(path.dirname(SCRIPT_PATH), "..");

const INCLUDED_DIRS = ["apps", "packages", "services", "crates", "tools"];

// Explicitly excluded paths (relative to repo root). We mostly avoid these by
// only scanning included source directories, but keep them here as belt+suspenders.
const EXCLUDED_ROOT_PATHS = new Set([
  "docs",
  "instructions",
  "mockups",
  "scratchpad.md",
  "handoff.md",
  "pnpm-lock.yaml",
]);

// These files document the policy and may include forbidden words intentionally.
const ALLOWLISTED_FILES = new Set([
  "AGENTS.md",
  "instructions/ai.md",
  "docs/05-ai-integration.md",
]);

const SKIP_DIR_NAMES = new Set([
  ".git",
  "node_modules",
  "dist",
  "coverage",
  "target",
  "build",
  ".turbo",
  ".pnpm-store",
  ".cache",
  ".vite",
  "playwright-report",
  "test-results",
  "security-report",
]);

// Keep this intentionally small: source code + config that can reintroduce deps.
const SCANNED_FILE_EXTENSIONS = new Set([
  ".js",
  ".jsx",
  ".ts",
  ".tsx",
  ".mjs",
  ".cjs",
  ".mts",
  ".cts",
  ".json",
  ".yaml",
  ".yml",
  ".toml",
  ".rs",
  ".py",
  ".go",
  ".java",
  ".kt",
  ".swift",
  ".cs",
  ".c",
  ".cc",
  ".cpp",
  ".h",
  ".hpp",
  ".proto",
  ".sh",
]);

const MAX_BYTES_TO_SCAN = 2 * 1024 * 1024; // 2 MiB guard against generated/binary blobs.

/**
 * @typedef {{ file: string, ruleId: string, message: string, line?: number, column?: number }} Violation
 */

/**
 * @typedef {{
 *   rootDir?: string,
 *   includedDirs?: string[],
 *   maxViolations?: number,
 * }} CheckOptions
 */

/**
 * @param {string} filePath
 * @param {string} content
 * @param {number} index
 */
function computeLineColumn(content, index) {
  // Line/col are for UX only; keep it cheap and simple.
  const upTo = content.slice(0, Math.max(0, index));
  const lines = upTo.split("\n");
  const line = lines.length;
  const column = lines[lines.length - 1]?.length + 1;
  return { line, column };
}

/**
 * Allowlisting for unit tests: tests may only mention forbidden patterns if they
 * are explicitly testing this guard script.
 *
 * @param {string} relativePath posix-ish relative path (we normalize separators)
 */
function isPolicyGuardTestFile(relativePath) {
  const base = path.basename(relativePath).toLowerCase();
  // Name-based allowlist so the rule is explicit and hard to "accidentally" hit.
  return base.includes("cursor-ai-policy") || base.includes("check-cursor-ai-policy");
}

/**
 * @param {string} relativePath
 */
function shouldExcludeRootRelativePath(relativePath) {
  const normalized = relativePath.split(path.sep).join("/");
  if (ALLOWLISTED_FILES.has(normalized)) return true;
  if (EXCLUDED_ROOT_PATHS.has(normalized)) return true;
  // Also exclude children of explicitly excluded root dirs.
  for (const excluded of ["docs", "instructions", "mockups"]) {
    if (normalized === excluded || normalized.startsWith(`${excluded}/`)) return true;
  }
  return false;
}

/**
 * @param {string} filePath absolute
 * @param {string} rootDir absolute
 */
function relativeToRoot(filePath, rootDir) {
  const rel = path.relative(rootDir, filePath);
  // Normalize separators for stable output across platforms.
  return rel.split(path.sep).join("/");
}

/**
 * @param {string} filePath absolute
 */
function shouldScanFile(filePath) {
  const ext = path.extname(filePath);
  if (ext && !SCANNED_FILE_EXTENSIONS.has(ext)) return false;
  // Files without extensions (rare) are ignored by default to reduce false
  // positives on license/readme-style blobs that can live inside packages.
  if (!ext) return false;
  return true;
}

/**
 * @param {string} relativePath normalized with forward slashes
 * @returns {boolean}
 */
function isTestFile(relativePath) {
  // Match `*.test.*` and `*.spec.*`.
  return /\.test\.[^.\\/]+$/i.test(relativePath) || /\.spec\.[^.\\/]+$/i.test(relativePath);
}

const CONTENT_SUBSTRING_RULES = [
  {
    id: "api.openai.com",
    needleLower: "api.openai.com",
    message: "Forbidden: direct OpenAI API endpoint reference (Cursor-only AI).",
  },
  {
    id: "anthropic",
    needleLower: "anthropic",
    message: "Forbidden: Anthropic integration reference (Cursor-only AI).",
  },
  {
    id: "ollama",
    needleLower: "ollama",
    message: "Forbidden: Ollama/local model integration reference (Cursor-only AI).",
  },
  {
    id: "formula:llm:",
    needleLower: "formula:llm:",
    message: "Forbidden: legacy/localStorage LLM key prefix (Cursor-only AI; no user keys).",
  },
  {
    id: "formula:openaiApiKey",
    needleLower: "formula:openaiapikey",
    message: "Forbidden: legacy OpenAI API key storage key (Cursor-only AI; no user keys).",
  },
  {
    id: "formula:aiCompletion:localModelEnabled",
    needleLower: "formula:aicompletion:localmodelenabled",
    message: "Forbidden: legacy local model toggle storage key (Cursor-only AI; no local models).",
  },
];

const OPENAI_IMPORT_REGEXES = [
  {
    id: "openai-import-from",
    re: /\b(?:import|export)\s+(?:type\s+)?[^"']*?\sfrom\s+["'][^"']*openai[^"']*["']/i,
    message: "Forbidden: OpenAI import specifier (Cursor-only AI).",
  },
  {
    id: "openai-import-side-effect",
    re: /\bimport\s+["'][^"']*openai[^"']*["']/i,
    message: "Forbidden: OpenAI import specifier (Cursor-only AI).",
  },
  {
    id: "openai-import-dynamic",
    re: /\bimport\(\s*["'][^"']*openai[^"']*["']\s*\)/i,
    message: "Forbidden: OpenAI dynamic import specifier (Cursor-only AI).",
  },
  {
    id: "openai-require",
    re: /\brequire\(\s*["'][^"']*openai[^"']*["']\s*\)/i,
    message: "Forbidden: OpenAI require() specifier (Cursor-only AI).",
  },
  {
    id: "openai-require-resolve",
    re: /\brequire\.resolve\(\s*["'][^"']*openai[^"']*["']\s*\)/i,
    message: "Forbidden: OpenAI require.resolve() specifier (Cursor-only AI).",
  },
];

/**
 * @param {string} absoluteDir
 * @param {string} rootDir
 * @param {(filePath: string) => Promise<void>} onFile
 */
async function walkDir(absoluteDir, rootDir, onFile) {
  /** @type {Array<{ dir: string }>} */
  const stack = [{ dir: absoluteDir }];
  while (stack.length) {
    const { dir } = /** @type {{dir: string}} */ (stack.pop());
    let entries;
    try {
      entries = await readdir(dir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      const rel = relativeToRoot(fullPath, rootDir);
      if (shouldExcludeRootRelativePath(rel)) continue;

      if (entry.isDirectory()) {
        if (SKIP_DIR_NAMES.has(entry.name) || entry.name.startsWith(".tmp")) continue;
        stack.push({ dir: fullPath });
        continue;
      }

      if (!entry.isFile()) continue;
      await onFile(fullPath);
    }
  }
}

/**
 * @param {CheckOptions} [options]
 * @returns {Promise<{ ok: boolean, violations: Violation[] }>}
 */
export async function checkCursorAiPolicy(options = {}) {
  const rootDir = options.rootDir ? path.resolve(options.rootDir) : DEFAULT_REPO_ROOT;
  const includedDirs = options.includedDirs ?? INCLUDED_DIRS;
  const maxViolations = options.maxViolations ?? 50;

  /** @type {Violation[]} */
  const violations = [];

  function record(v) {
    violations.push(v);
  }

  /**
   * @param {string} filePath absolute
   */
  async function scanFile(filePath) {
    if (violations.length >= maxViolations) return;

    const rel = relativeToRoot(filePath, rootDir);
    if (shouldExcludeRootRelativePath(rel)) return;

    const relLower = rel.toLowerCase();

    // Forbidden filename/path patterns.
    if (relLower.includes("openai")) {
      record({ file: rel, ruleId: "path-openai", message: "Forbidden: `openai` in file path (Cursor-only AI)." });
      return;
    }
    if (relLower.includes("anthropic")) {
      record({
        file: rel,
        ruleId: "path-anthropic",
        message: "Forbidden: `anthropic` in file path (Cursor-only AI).",
      });
      return;
    }
    if (relLower.includes("ollama")) {
      record({ file: rel, ruleId: "path-ollama", message: "Forbidden: `ollama` in file path (Cursor-only AI)." });
      return;
    }

    if (!shouldScanFile(filePath)) return;

    let st;
    try {
      st = await stat(filePath);
    } catch {
      return;
    }
    if (!st.isFile()) return;
    if (st.size > MAX_BYTES_TO_SCAN) return;

    let content;
    try {
      content = await readFile(filePath, "utf8");
    } catch {
      return;
    }

    const isTest = isTestFile(rel);
    const allowForbiddenInThisTest = isTest && isPolicyGuardTestFile(rel);

    // For test files, keep the exception narrow and explicit.
    // We still *scan* them (so unrelated tests can't mention forbidden providers),
    // but allow guard-specific tests to include those strings.
    if (allowForbiddenInThisTest) return;

    const contentLower = content.toLowerCase();
    for (const rule of CONTENT_SUBSTRING_RULES) {
      const idx = contentLower.indexOf(rule.needleLower);
      if (idx === -1) continue;
      const { line, column } = computeLineColumn(content, idx);
      record({ file: rel, ruleId: rule.id, message: rule.message, line, column });
      return;
    }

    // OpenAI is checked more narrowly (import specifiers) to reduce accidental matches.
    for (const rule of OPENAI_IMPORT_REGEXES) {
      const m = rule.re.exec(content);
      if (!m) continue;
      const idx = m.index ?? contentLower.indexOf("openai");
      const { line, column } = computeLineColumn(content, idx);
      record({ file: rel, ruleId: rule.id, message: rule.message, line, column });
      return;
    }
  }

  const dirsToScan = [];
  for (const dir of includedDirs) {
    const abs = path.join(rootDir, dir);
    try {
      const st = await stat(abs);
      if (st.isDirectory()) dirsToScan.push(abs);
    } catch {
      // ignore missing
    }
  }

  for (const dir of dirsToScan) {
    await walkDir(dir, rootDir, scanFile);
    if (violations.length >= maxViolations) break;
  }

  return { ok: violations.length === 0, violations };
}

function formatViolations(violations) {
  return violations
    .map((v) => {
      const loc = v.line ? `:${v.line}:${v.column}` : "";
      return `- ${v.file}${loc} [${v.ruleId}] ${v.message}`;
    })
    .join(os.EOL);
}

async function main() {
  const args = process.argv.slice(2);
  let rootDir = null;
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--help" || arg === "-h") {
      console.log(`Usage: node ${path.relative(process.cwd(), SCRIPT_PATH)} [--root <dir>]`);
      process.exit(0);
    }
    if (arg === "--root") {
      rootDir = args[i + 1] ? String(args[i + 1]) : "";
      i++;
      continue;
    }
    if (arg.startsWith("--root=")) {
      rootDir = arg.slice("--root=".length);
      continue;
    }
  }

  const result = await checkCursorAiPolicy({ rootDir: rootDir ?? undefined });
  if (result.ok) {
    process.exit(0);
  }

  console.error("Cursor-only AI policy violation(s) found:");
  console.error(formatViolations(result.violations));
  process.exit(1);
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(SCRIPT_PATH)) {
  await main();
}
