import { describe, it } from "vitest";

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils";

const SRC_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".mts", ".cts"]);
const WINDOW_DIALOG_RES = [
  // Direct calls: window.confirm / window.confirm?.; whitespace around the dot allowed.
  /\bwindow\s*\.\s*(confirm|alert|prompt)\s*(?:\?\.)?\s*\(/,
  // Optional chaining on window: window?.confirm / window?.confirm?.
  /\bwindow\s*\?\.\s*(confirm|alert|prompt)\s*(?:\?\.)?\s*\(/,
  // Bracket access: window["confirm"] / window["confirm"]?.
  /\bwindow\s*\[\s*['"](confirm|alert|prompt)['"]\s*\]\s*(?:\?\.)?\s*\(/,
  // Optional chaining + bracket access: window?.["confirm"]
  /\bwindow\s*\?\.\s*\[\s*['"](confirm|alert|prompt)['"]\s*\]\s*(?:\?\.)?\s*\(/,
  // Direct calls via other global roots: globalThis.confirm / self.confirm / etc.
  /\b(?:globalThis|self)\s*\.\s*(confirm|alert|prompt)\s*(?:\?\.)?\s*\(/,
  /\b(?:globalThis|self)\s*\?\.\s*(confirm|alert|prompt)\s*(?:\?\.)?\s*\(/,
  /\b(?:globalThis|self)\s*\[\s*['"](confirm|alert|prompt)['"]\s*\]\s*(?:\?\.)?\s*\(/,
  /\b(?:globalThis|self)\s*\?\.\s*\[\s*['"](confirm|alert|prompt)['"]\s*\]\s*(?:\?\.)?\s*\(/,
];
const WINDOW_DIALOG_RE_SOURCE = WINDOW_DIALOG_RES.map((re) => `(?:${re.source})`).join("|");
const WINDOW_DIALOG_RE = new RegExp(WINDOW_DIALOG_RE_SOURCE);

async function collectSourceFiles(dir: string): Promise<string[]> {
  const out: string[] = [];
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const abs = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await collectSourceFiles(abs)));
      continue;
    }
    if (!entry.isFile()) continue;
    if (!SOURCE_EXTS.has(path.extname(entry.name))) continue;
    out.push(abs);
  }
  return out;
}

function findViolationsInFileContent(content: string, relPath: string): string[] {
  // Fast-path: the banned patterns require:
  // - a global root (`window` / `globalThis` / `self`), and
  // - one of the dialog API names (`confirm` / `alert` / `prompt`).
  //
  // Avoid running the heavier comment-stripping scan for files that cannot possibly match.
  //
  // Use word boundaries so we don't scan files with unrelated substrings like `windowing` or
  // `confirmation`.
  if (!/\b(?:confirm|alert|prompt)\b/.test(content)) return [];
  if (!/\b(?:window|globalThis|self)\b/.test(content)) return [];

  // Strip comments so commented-out `window.alert(...)` wiring cannot satisfy or fail assertions.
  const stripped = stripComments(content);

  // Only compute line/column context when there is a match. The common case is that
  // there are *no* violations, so avoid splitting every file into lines (which gets
  // expensive as the codebase grows).
  if (!WINDOW_DIALOG_RE.test(stripped)) return [];

  const violations: string[] = [];
  const re = new RegExp(WINDOW_DIALOG_RE_SOURCE, "g");
  const rawLines = content.split(/\r?\n/);
  re.lastIndex = 0;
  let match: RegExpExecArray | null = null;
  while ((match = re.exec(stripped))) {
    const idx = match.index;
    const lineStart = stripped.lastIndexOf("\n", idx) + 1;
    // Best-effort line number: only needed when reporting violations, so a simple
    // split is fine here.
    const lineNumber = stripped.slice(0, lineStart).split(/\r?\n/).length;
    const lineText = rawLines[lineNumber - 1] ?? "";
    violations.push(`${relPath}:${lineNumber}: ${lineText.trim()}`);
  }
  return violations;
}

describe("desktop/src should not use blocking browser dialogs", () => {
  // This is a source scan over the entire desktop renderer tree and can take
  // longer than Vitest's default 30s timeout in CI / constrained environments.
  it("does not call window.alert/confirm/prompt anywhere under apps/desktop/src", async () => {
    const files = await collectSourceFiles(SRC_ROOT);
    const violations: string[] = [];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      const content = await readFile(absPath, "utf8");
      violations.push(...findViolationsInFileContent(content, relPath));
    }

    if (violations.length > 0) {
      throw new Error(`Found blocking browser dialogs in desktop renderer code:\n${violations.join("\n")}`);
    }
  }, 60_000);
});
