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

// Keep this in sync with `tauri/apiAccessGuard.vitest.ts` so we cover common access patterns
// (dot access, optional chaining, and bracket access to the global itself).
const TAURI_GLOBAL_REF_RE_SOURCE =
  "(?:\\(\\s*(?:globalThis|window|self)\\s+as\\s+any\\s*\\)\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b|(?:globalThis|window|self)\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b|__TAURI__\\b|\\(\\s*(?:globalThis|window|self)\\s+as\\s+any\\s*\\)\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]|(?:globalThis|window|self)\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\])";

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

function collectTauriAliases(content: string): Set<string> {
  const tauriRoots = new Set<string>();

  // Capture common aliasing patterns like:
  //   const tauri = (globalThis as any).__TAURI__;
  //   let tauri = globalThis.__TAURI__ ?? null;
  //
  // NOTE: This targets only direct aliases to the root `__TAURI__` object (not nested properties).
  const tauriRootAssignRe =
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__\b|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__\b|__TAURI__\b|\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]|(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\])(?!(?:\s*(?:\?\.|\.|\[)))/g;

  let match: RegExpExecArray | null = null;
  while ((match = tauriRootAssignRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriRoots.add(name);
    if (match[0].length === 0) tauriRootAssignRe.lastIndex += 1;
  }

  return tauriRoots;
}

function collectTauriCoreAliases(content: string, tauriRoots: Set<string>): Set<string> {
  const coreAliases = new Set<string>();

  // Alias patterns like:
  //   const core = __TAURI__.core;
  //   const core = globalThis.__TAURI__?.core;
  //   const core = tauri["core"];
  const coreAssignDotRe = new RegExp(
    `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\s*(?:\\?\\.|\\.)\\s*core\\b`,
    "g",
  );
  const coreAssignBracketRe = new RegExp(
    `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]core['\\\"]\\s*\\]`,
    "g",
  );

  for (const re of [coreAssignDotRe, coreAssignBracketRe]) {
    let match: RegExpExecArray | null = null;
    while ((match = re.exec(content)) != null) {
      const name = match[1];
      if (name) coreAliases.add(name);
      if (match[0].length === 0) re.lastIndex += 1;
    }
  }

  // Destructuring patterns like:
  //   const { core } = __TAURI__;
  //   const { core: tauriCore } = tauri;
  const coreDestructureDirectRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bcore\\b(?!\\s*:)` +
      `[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\b`,
    "g",
  );
  const coreDestructureRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bcore\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\b`,
    "g",
  );

  if (coreDestructureDirectRe.test(content)) coreAliases.add("core");

  let match: RegExpExecArray | null = null;
  while ((match = coreDestructureRenameRe.exec(content)) != null) {
    const name = match[1];
    if (name) coreAliases.add(name);
    if (match[0].length === 0) coreDestructureRenameRe.lastIndex += 1;
  }

  for (const root of tauriRoots) {
    const r = escapeRegExp(root);

    const coreAssignFromAliasDotRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*core\\b`,
      "g",
    );
    const coreAssignFromAliasBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]core['\\\"]\\s*\\]`,
      "g",
    );

    for (const re of [coreAssignFromAliasDotRe, coreAssignFromAliasBracketRe]) {
      while ((match = re.exec(content)) != null) {
        const name = match[1];
        if (name) coreAliases.add(name);
        if (match[0].length === 0) re.lastIndex += 1;
      }
    }

    // Destructuring from an alias: `const { core } = tauri;`
    const coreDestructureFromAliasDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bcore\\b(?!\\s*:)[^}]*\\}\\s*=\\s*${r}\\b`,
      "g",
    );
    if (coreDestructureFromAliasDirectRe.test(content)) coreAliases.add("core");

    const coreDestructureFromAliasRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bcore\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${r}\\b`,
      "g",
    );
    while ((match = coreDestructureFromAliasRenameRe.exec(content)) != null) {
      const name = match[1];
      if (name) coreAliases.add(name);
      if (match[0].length === 0) coreDestructureFromAliasRenameRe.lastIndex += 1;
    }
  }

  return coreAliases;
}

function buildBannedResForTauriAlias(root: string): RegExp[] {
  const r = escapeRegExp(root);
  return [
    // Direct access via alias: tauri.core.invoke / tauri?.core?.invoke / etc.
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*core\\s*(?:\\?\\.|\\.)\\s*invoke\\b`),
    // Mixed bracket/dot access: tauri["core"].invoke / tauri.core["invoke"].
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]core['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*invoke\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*core\\s*(?:\\?\\.)?\\s*\\[\\s*['"]invoke['"]\\s*\\]`),
    // Bracket access: tauri["core"]["invoke"] / tauri?.["core"]?.["invoke"]
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]core['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]invoke['"]\\s*\\]`),
    // Destructuring patterns: `const { invoke } = tauri.core;` / `const { core: { invoke } } = tauri;`
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*core\\b`,
    ),
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]core['"]\\s*\\]`,
    ),
    new RegExp(`\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\bcore\\b[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\b`),
  ];
}

function buildBannedResForTauriCoreAlias(core: string): RegExp[] {
  const r = escapeRegExp(core);
  return [
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*invoke\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]invoke['"]\\s*\\]`),
    // Destructuring patterns: `const { invoke } = core;`
    new RegExp(`\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\b`),
  ];
}

describe("tauri/invoke guardrails", () => {
  it("does not access __TAURI__.core.invoke outside src/tauri helpers", async () => {
    const files = await collectSourceFiles(SRC_ROOT);
    const violations = new Set<string>();

    // Keep this intentionally scoped to *direct* core.invoke property access so we don't ban other
    // legitimate `__TAURI__` uses (plugins, etc).
    const bannedRes: RegExp[] = [
        // __TAURI__.core.invoke / __TAURI__?.core?.invoke / mixed optional chaining.
        /\b__TAURI__\s*(?:\?\.|\.)\s*core\s*(?:\?\.|\.)\s*invoke\b/,
        // Mixed bracket/dot variants: __TAURI__["core"].invoke / __TAURI__.core["invoke"].
        /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]core['"]\s*\]\s*(?:\?\.|\.)\s*invoke\b/,
        /\b__TAURI__\s*(?:\?\.|\.)\s*core\s*(?:\?\.)?\s*\[\s*['"]invoke['"]\s*\]/,
        // Bracket access variants: __TAURI__["core"]["invoke"] / __TAURI__?.["core"]?.["invoke"].
        /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]core['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]invoke['"]\s*\]/,
        // Bracket access to the __TAURI__ global itself (e.g. globalThis["__TAURI__"].core.invoke).
        /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*core\s*(?:\?\.|\.)\s*invoke\b/,
        // Mixed bracket/dot access to globals: globalThis["__TAURI__"]["core"].invoke / globalThis["__TAURI__"].core["invoke"].
        /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]core['"]\s*\]\s*(?:\?\.|\.)\s*invoke\b/,
        /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*core\s*(?:\?\.)?\s*\[\s*['"]invoke['"]\s*\]/,
        /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]core['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]invoke['"]\s*\]/,
        // Destructuring patterns: `const { invoke } = __TAURI__.core;` / `const { core: { invoke } } = __TAURI__;`
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\s*(?:\\?\\.|\\.)\\s*core\\b`,
        ),
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]core['"]\\s*\\]`,
        ),
        new RegExp(`\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\bcore\\b[\\s\\S]*?\\binvoke\\b[\\s\\S]*?\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}`),
      ];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      if (isTestFile(relPath)) continue;

      const normalized = relPath.replace(/\\/g, "/");
      // The canonical locations for core.invoke access.
      if (normalized === "tauri/api.ts" || normalized === "tauri/api.js") continue;
      if (normalized === "tauri/invoke.js" || normalized === "tauri/invoke.ts") continue;

      const content = await readFile(absPath, "utf8");

      const matches = (re: RegExp) => re.test(content);
      if (bannedRes.some(matches)) {
        violations.add(normalized);
        continue;
      }

      const aliases = collectTauriAliases(content);
      const coreAliases = collectTauriCoreAliases(content, aliases);
      if (aliases.size === 0 && coreAliases.size === 0) continue;

      const aliasRes: RegExp[] = [];
      for (const alias of aliases) {
        aliasRes.push(...buildBannedResForTauriAlias(alias));
      }
      const coreAliasRes: RegExp[] = [];
      for (const alias of coreAliases) {
        coreAliasRes.push(...buildBannedResForTauriCoreAlias(alias));
      }
      if ([...aliasRes, ...coreAliasRes].some(matches)) {
        violations.add(normalized);
      }
    }

    if (violations.size > 0) {
      throw new Error(
        "Direct __TAURI__.core.invoke access is not allowed outside `src/tauri` helpers.\n\nViolations:\n" +
          Array.from(violations)
            .sort()
            .map((p) => `- ${p}`)
            .join("\n"),
      );
    }
  });
});
