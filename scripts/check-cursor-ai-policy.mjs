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

// Default scan roots. We intentionally include:
// - top-level `test/` + `tests/` so the "no provider names in unit tests" rule
//   applies to both workspace tests and the repo's node:test suites.
// - `shared/` and `extensions/` so common libs + first-party extension bundles
//   can't accidentally reintroduce provider-specific deps/config.
// - `scripts/` and `python/` so CI/util scripts can't silently add provider
//   integrations outside the main app/package trees.
// - `fixtures/` so fixture generators and committed test assets can't hide
//   provider integrations or config.
const INCLUDED_DIRS = [
  "apps",
  "packages",
  "services",
  "crates",
  "tools",
  "shared",
  "extensions",
  "fixtures",
  "patches",
  "security",
  ".github",
  ".cargo",
  ".vscode",
  ".devcontainer",
  "scripts",
  "python",
  "test",
  "tests",
];

// Explicitly excluded paths (relative to repo root). We mostly avoid these by
// only scanning included source directories, but keep them here as belt+suspenders.
const EXCLUDED_ROOT_PATHS = new Set([
  "docs",
  "instructions",
  "mockups",
  "scratchpad.md",
  "handoff.md",
]);

// These files document the policy (or implement this guard) and may include
// forbidden words intentionally.
const ALLOWLISTED_FILES = new Set([
  "AGENTS.md",
  "instructions/ai.md",
  "docs/05-ai-integration.md",
  "scripts/check-cursor-ai-policy.mjs",
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
  ".ps1",
  ".css",
  ".md",
  ".html",
  ".snap",
  ".patch",
  ".diff",
  ".json",
  ".jsonl",
  ".yaml",
  ".yml",
  ".plist",
  ".toml",
  ".lock",
  ".sql",
  ".xml",
  ".tsv",
  ".csv",
  ".txt",
  ".bas",
  ".m",
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
  ".wgsl",
  ".b64",
  ".base64",
  ".pem",
  ".key",
  ".crt",
  ".sh",
]);

const MAX_BYTES_TO_SCAN = 2 * 1024 * 1024; // 2 MiB guard against generated/binary blobs.

// Config formats where any mention of provider identifiers is almost certainly a
// direct dependency/config regression (not an incidental string in code/comments).
const CONFIG_FILE_EXTENSIONS = new Set([".json", ".toml", ".yaml", ".yml", ".lock"]);

// Some important config/build files have no extension. Scan these by basename so
// they can't be used to smuggle provider integrations into the repo.
//
// Keep this list intentionally small; prefer adding explicit basenames over
// scanning *all* extensionless files (which increases false positives).
const SCANNED_BASENAMES_WITHOUT_EXTENSION = new Set([
  "dockerfile",
  "makefile",
  "config",
  "license",
  "notice",
  ".gitkeep",
  ".gitignore",
  ".gitattributes",
  ".dockerignore",
  ".npmrc",
  ".yarnrc",
  ".editorconfig",
]);

// Some config files embed their "extension" into the basename (notably `.env`,
// `.env.local`, `.env.production`, etc). Scan them by prefix so they can't be
// used to stash provider keys in a committed env file.
const SCANNED_BASENAME_PREFIXES = [".env", "dockerfile", "makefile"];

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
  const extLower = ext ? ext.toLowerCase() : "";
  const baseLower = path.basename(filePath).toLowerCase();
  if (SCANNED_BASENAME_PREFIXES.some((prefix) => baseLower.startsWith(prefix))) return true;
  if (extLower && !SCANNED_FILE_EXTENSIONS.has(extLower)) return false;
  // Files without extensions (rare) are ignored by default to reduce false
  // positives on license/readme-style blobs that can live inside packages.
  if (!extLower) {
    return SCANNED_BASENAMES_WITHOUT_EXTENSION.has(baseLower);
  }
  return true;
}

/**
 * @param {string} relativePath normalized with forward slashes
 * @returns {boolean}
 */
function isTestFile(relativePath) {
  // Match `*.test.*`, `*.spec.*`, and `*.vitest.*`.
  return (
    /\.test\.[^.\\/]+$/i.test(relativePath) ||
    /\.spec\.[^.\\/]+$/i.test(relativePath) ||
    /\.vitest\.[^.\\/]+$/i.test(relativePath)
  );
}

const CONTENT_SUBSTRING_RULES = [
  {
    id: "api.openai.com",
    needleLower: "api.openai.com",
    message: "Forbidden: direct OpenAI API endpoint reference (Cursor-only AI).",
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
  {
    id: "openai",
    needleLower: "openai",
    message: "Forbidden: OpenAI integration reference (Cursor-only AI).",
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
];

const OPENAI_IN_TEST_RULE = {
  id: "openai-in-test",
  needleLower: "openai",
  message:
    "Forbidden: OpenAI references are not allowed in unit tests (Cursor-only AI). Use generic placeholders or test fixtures that avoid provider names.",
};

const OPENAI_IN_CONFIG_RULE = {
  id: "openai-in-config",
  needleLower: "openai",
  message: "Forbidden: OpenAI references are not allowed in config/dependency files (Cursor-only AI).",
};

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
        if (entry.name.startsWith(".tmp")) continue;
        if (SKIP_DIR_NAMES.has(entry.name)) {
          // Extensions ship their built bundles (under `extensions/**/dist/`) so
          // integration tests and marketplace packaging can run without an extra
          // build step. Scan those committed dist assets for policy violations.
          const normalized = rel.split(path.sep).join("/");
          if (!(entry.name === "dist" && normalized.startsWith("extensions/"))) continue;
        }
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
    const ext = path.extname(relLower);

    // For config/dependency files, any mention of provider identifiers is a hard fail.
    if (CONFIG_FILE_EXTENSIONS.has(ext)) {
      const idx = contentLower.indexOf(OPENAI_IN_CONFIG_RULE.needleLower);
      if (idx !== -1) {
        const { line, column } = computeLineColumn(content, idx);
        record({
          file: rel,
          ruleId: OPENAI_IN_CONFIG_RULE.id,
          message: OPENAI_IN_CONFIG_RULE.message,
          line,
          column,
        });
        return;
      }
    }

    // Unit tests should not mention provider names at all (unless they are tests for
    // this guard). This is stricter than the import-specifier check and helps
    // prevent regressions via "stringly-typed" provider selection.
    if (isTest) {
      const idx = contentLower.indexOf(OPENAI_IN_TEST_RULE.needleLower);
      if (idx !== -1) {
        const { line, column } = computeLineColumn(content, idx);
        record({ file: rel, ruleId: OPENAI_IN_TEST_RULE.id, message: OPENAI_IN_TEST_RULE.message, line, column });
        return;
      }
    }

    for (const rule of CONTENT_SUBSTRING_RULES) {
      const idx = contentLower.indexOf(rule.needleLower);
      if (idx === -1) continue;
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

  // Also scan root-level config files (package.json, Cargo.toml, etc). Those can
  // reintroduce forbidden dependencies without touching the main code trees.
  if (violations.length < maxViolations) {
    try {
      const rootEntries = await readdir(rootDir, { withFileTypes: true });
      for (const entry of rootEntries) {
        if (violations.length >= maxViolations) break;
        if (!entry.isFile()) continue;
        await scanFile(path.join(rootDir, entry.name));
      }
    } catch {
      // ignore
    }
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
