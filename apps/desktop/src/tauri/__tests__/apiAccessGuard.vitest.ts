import { describe, it } from "vitest";

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const TAURI_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC_ROOT = path.resolve(TAURI_DIR, "..");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".mts", ".cts"]);

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

function isTestFile(relPath: string): boolean {
  const normalized = relPath.replace(/\\/g, "/");
  if (normalized.includes("/__tests__/")) return true;
  if (/\.(test|vitest)\.[^./]+$/.test(normalized)) return true;
  return false;
}

describe("tauri/api guardrails", () => {
  it("does not access __TAURI__.event / __TAURI__.window / __TAURI__.dialog.* outside src/tauri/api", async () => {
    const files = await collectSourceFiles(SRC_ROOT);
    const violations = new Set<string>();

    // Keep these regexes intentionally narrow so we don't block other (non-event/window)
    // uses of `__TAURI__` in the renderer (e.g. core.invoke, notifications, etc).
    // These regexes are applied against the full file contents (not line-by-line) so we also
    // catch multi-line chains like:
    //   (globalThis as any).__TAURI__?.dialog
    //     ?.open(...)
    const bannedRes: RegExp[] = [
      // Event API access (listen/emit) should go through getTauriEventApiOr{Null,Throw}.
      /\b__TAURI__\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*event\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*event\b/,
      // Bracket access variants: __TAURI__["event"] / __TAURI__?.["event"] / __TAURI__["plugin"]["event"].
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,

      // Window API access should go through getTauriWindowHandleOr{Null,Throw} or hasTauriWindow* helpers.
      /\b__TAURI__\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,

      // Dialog API access should go through tauri/api helpers (or `nativeDialogs` where appropriate).
      /\b__TAURI__\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*\.\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
    ];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      if (isTestFile(relPath)) continue;

      const normalized = relPath.replace(/\\/g, "/");
      if (normalized === "tauri/api.ts" || normalized === "tauri/api.js") continue;

      const content = await readFile(absPath, "utf8");
      const lines = content.split(/\r?\n/);

      for (const re of bannedRes) {
        const globalRe = new RegExp(re.source, re.flags.includes("g") ? re.flags : `${re.flags}g`);
        let match: RegExpExecArray | null = null;
        while ((match = globalRe.exec(content)) != null) {
          const start = match.index;
          const lineNumber = content.slice(0, start).split(/\r?\n/).length;
          const line = lines[lineNumber - 1] ?? "";
          violations.add(`${relPath}:${lineNumber}: ${line.trim()}`);

          // Avoid infinite loops on zero-length matches.
          if (match[0].length === 0) globalRe.lastIndex += 1;
        }
      }
    }

    if (violations.size > 0) {
      throw new Error(
        "Found direct __TAURI__ dialog/window/event access outside src/tauri/api:\n" + [...violations].join("\n"),
      );
    }
  });
});
