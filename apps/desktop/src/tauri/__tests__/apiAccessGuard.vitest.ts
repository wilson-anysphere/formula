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
  it("does not access __TAURI__.event / __TAURI__.window / dialog.open/save outside src/tauri/api", async () => {
    const files = await collectSourceFiles(SRC_ROOT);
    const violations: string[] = [];

    // Keep these regexes intentionally narrow so we don't block other (non-event/window)
    // uses of `__TAURI__` in the renderer (e.g. core.invoke, notifications, etc).
    const bannedLineRes: RegExp[] = [
      // Event API access (listen/emit) should go through getTauriEventApiOr{Null,Throw}.
      /\b__TAURI__\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*event\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*event\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*event\b/,

      // Window API access should go through getTauriWindowHandleOr{Null,Throw} or hasTauriWindow* helpers.
      /\b__TAURI__\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*window\b/,

      // Dialog open/save should go through getTauriDialogOr{Null,Throw}. (Confirm/alert are handled
      // separately by nativeDialogs and are intentionally not included here.)
      /\b__TAURI__\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*(open|save)\b/,
      /\b__TAURI__\s*\.\s*dialog\s*(?:\?\.)\s*(open|save)\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*(open|save)\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*(open|save)\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*(open|save)\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*(open|save)\b/,

      // Avoid ad-hoc checks for Tauri confirm dialogs; use nativeDialogs or tauri/api helpers.
      /\b__TAURI__\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*confirm\b/,
      /\b__TAURI__\s*\.\s*dialog\s*(?:\?\.)\s*confirm\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugin\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*confirm\b/,
      /\b__TAURI__\s*\.\s*plugin\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*confirm\b/,
      /\b__TAURI__\s*(?:\?\.)\s*plugins\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*confirm\b/,
      /\b__TAURI__\s*\.\s*plugins\s*(?:\?\.)\s*dialog\s*(?:\?\.)\s*confirm\b/,
    ];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      if (isTestFile(relPath)) continue;

      const normalized = relPath.replace(/\\/g, "/");
      if (normalized === "tauri/api.ts" || normalized === "tauri/api.js") continue;

      const content = await readFile(absPath, "utf8");
      const lines = content.split(/\r?\n/);
      for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i] ?? "";
        if (bannedLineRes.some((re) => re.test(line))) {
          violations.push(`${relPath}:${i + 1}: ${line.trim()}`);
        }
      }
    }

    if (violations.length > 0) {
      throw new Error(
        "Found direct __TAURI__ dialog/window/event access outside src/tauri/api:\n" + violations.join("\n"),
      );
    }
  });
});
