import { describe, it } from "vitest";

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

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

describe("desktop/src should not use blocking browser dialogs", () => {
  it("does not call window.alert/confirm/prompt anywhere under apps/desktop/src", async () => {
    const files = await collectSourceFiles(SRC_ROOT);
    const violations: string[] = [];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      const content = await readFile(absPath, "utf8");
      const lines = content.split(/\r?\n/);
      for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i] ?? "";
        if (WINDOW_DIALOG_RES.some((re) => re.test(line))) {
          violations.push(`${relPath}:${i + 1}: ${line.trim()}`);
        }
      }
    }

    if (violations.length > 0) {
      throw new Error(`Found blocking browser dialogs in desktop renderer code:\n${violations.join("\n")}`);
    }
  });
});
