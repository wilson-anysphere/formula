import { describe, it } from "vitest";

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const TAURI_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC_ROOT = path.resolve(TAURI_DIR, "..");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".mts", ".cts"]);

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

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

type TauriAliasSets = {
  tauriRoots: Set<string>;
  tauriPluginRoots: Set<string>;
  tauriPluginsRoots: Set<string>;
};

function collectTauriAliases(content: string): TauriAliasSets {
  const tauriRoots = new Set<string>();
  const tauriPluginRoots = new Set<string>();
  const tauriPluginsRoots = new Set<string>();

  // Capture common aliasing patterns like:
  //   const tauri = (globalThis as any).__TAURI__;
  //   let tauri = globalThis.__TAURI__ ?? null;
  //
  // NOTE: This intentionally only targets direct aliases to the root `__TAURI__` object (not
  // nested properties like `__TAURI__.core.invoke`), so we can then flag `tauri.dialog` /
  // `tauri.window` / `tauri.event` access even when the file doesn't mention `__TAURI__` again.
  const tauriRootAssignRe =
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__|__TAURI__)\b(?!\s*(?:\?\.|\.|\[))/g;

  let match: RegExpExecArray | null = null;
  while ((match = tauriRootAssignRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriRoots.add(name);
    if (match[0].length === 0) tauriRootAssignRe.lastIndex += 1;
  }

  // Capture aliases to the plugin container objects:
  //   const plugin = (globalThis as any).__TAURI__?.plugin;
  const tauriPluginAssignRe =
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__|__TAURI__)\s*(?:\?\.|\.)\s*(plugin|plugins)\b(?!\s*(?:\?\.|\.|\[))/g;

  while ((match = tauriPluginAssignRe.exec(content)) != null) {
    const name = match[1];
    const which = match[2];
    if (name && which === "plugin") tauriPluginRoots.add(name);
    if (name && which === "plugins") tauriPluginsRoots.add(name);
    if (match[0].length === 0) tauriPluginAssignRe.lastIndex += 1;
  }

  return { tauriRoots, tauriPluginRoots, tauriPluginsRoots };
}

function buildBannedResForTauriAlias(root: string): RegExp[] {
  const r = escapeRegExp(root);
  return [
    // Direct root access: tauri.event / tauri.window / tauri.dialog.
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),

    // Plugin container variants.
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
  ];
}

function buildBannedResForTauriPluginAlias(root: string): RegExp[] {
  const r = escapeRegExp(root);
  return [
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
  ];
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
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*event\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*event\b/,
      // Bracket access variants: __TAURI__["event"] / __TAURI__?.["event"] / __TAURI__["plugin"]["event"].
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,

      // Window API access should go through getTauriWindowHandleOr{Null,Throw} or hasTauriWindow* helpers.
      /\b__TAURI__\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*window\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,

      // Dialog API access should go through tauri/api helpers (or `nativeDialogs` where appropriate).
      /\b__TAURI__\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*\.\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*dialog\b/,
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

      const aliasRes: RegExp[] = [];
      const aliases = collectTauriAliases(content);
      for (const alias of aliases.tauriRoots) {
        aliasRes.push(...buildBannedResForTauriAlias(alias));
      }
      for (const alias of aliases.tauriPluginRoots) {
        aliasRes.push(...buildBannedResForTauriPluginAlias(alias));
      }
      for (const alias of aliases.tauriPluginsRoots) {
        aliasRes.push(...buildBannedResForTauriPluginAlias(alias));
      }

      for (const re of [...bannedRes, ...aliasRes]) {
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
