import { describe, it } from "vitest";

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

const TAURI_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC_ROOT = path.resolve(TAURI_DIR, "..");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs", ".mts", ".cts"]);

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Keep this in sync with `tauri/invokeAccessGuard.vitest.ts` so we cover common access patterns
// (dot access, optional chaining, and bracket access to the global itself).
const TAURI_GLOBAL_REF_RE_SOURCE =
  "(?:\\(\\s*(?:globalThis|window|self)\\s+as\\s+any\\s*\\)\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b|(?:globalThis|window|self)\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b|__TAURI__\\b|\\(\\s*(?:globalThis|window|self)\\s+as\\s+any\\s*\\)\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]|(?:globalThis|window|self)\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\])";

const GLOBAL_OBJECT_REF_RE_SOURCE =
  "(?:\\(\\s*(?:globalThis|window|self)\\s+as\\s+any\\s*\\)|(?:globalThis|window|self)(?:\\s+as\\s+any)?)";

function buildAnyCastableRefSource(escapedIdentifier: string): string {
  // Matches either:
  // - `(ident as any)` (starts with `(`, so cannot be preceded by `\\b`)
  // - `ident` / `ident as any` (word-boundary delimited)
  return `(?:\\(\\s*${escapedIdentifier}\\s+as\\s+any\\s*\\)|\\b${escapedIdentifier}\\b(?:\\s+as\\s+any)?)`;
}

function collectGlobalObjectAliases(content: string): Set<string> {
  const globalRoots = new Set<string>();

  // Capture patterns like:
  //   const g = globalThis;
  //   const w = window;
  //   const g = globalThis as any;
  //   const g = (globalThis as any);
  const globalAssignRe = new RegExp(
    `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}(?!\\s*(?:\\?\\.|\\.|\\[))`,
    "g",
  );

  let match: RegExpExecArray | null = null;
  while ((match = globalAssignRe.exec(content)) != null) {
    const name = match[1];
    if (name) globalRoots.add(name);
    if (match[0].length === 0) globalAssignRe.lastIndex += 1;
  }

  return globalRoots;
}

function collectTauriAliasesFromGlobalAliases(content: string, globalAliases: Set<string>): TauriAliasSets {
  const tauriRoots = new Set<string>();
  const tauriPluginRoots = new Set<string>();
  const tauriPluginsRoots = new Set<string>();

  if (globalAliases.size === 0) return { tauriRoots, tauriPluginRoots, tauriPluginsRoots };

  for (const globalAlias of globalAliases) {
    const r = escapeRegExp(globalAlias);
    const g = buildAnyCastableRefSource(r);

    const assignDotRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
      "g",
    );
    const assignBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]`,
      "g",
    );
    const destructureRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${g}`,
      "g",
    );
    const destructureRenameQuotedRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?['\\\"]__TAURI__['\\\"]\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${g}`,
      "g",
    );

    for (const re of [assignDotRe, assignBracketRe, destructureRenameRe, destructureRenameQuotedRe]) {
      let match: RegExpExecArray | null = null;
      while ((match = re.exec(content)) != null) {
        const name = match[1];
        if (name) tauriRoots.add(name);
        if (match[0].length === 0) re.lastIndex += 1;
      }
    }

    const pluginAssignRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.|\\.)\\s*(plugin|plugins)\\b`,
      "g",
    );
    const pluginAssignBracketTauriRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*(plugin|plugins)\\b`,
      "g",
    );
    for (const re of [pluginAssignRe, pluginAssignBracketTauriRe]) {
      let match: RegExpExecArray | null = null;
      while ((match = re.exec(content)) != null) {
        const name = match[1];
        const which = match[2];
        if (name && which === "plugin") tauriPluginRoots.add(name);
        if (name && which === "plugins") tauriPluginsRoots.add(name);
        if (match[0].length === 0) re.lastIndex += 1;
      }
    }

    // Destructuring plugin containers from g.__TAURI__ / g["__TAURI__"].
    const pluginDestructureDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b(?!\\s*:)` +
        `[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
    );
    const pluginsDestructureDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b(?!\\s*:)` +
        `[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
    );
    const pluginDestructureDirectBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b(?!\\s*:)` +
        `[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]`,
    );
    const pluginsDestructureDirectBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b(?!\\s*:)` +
        `[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]`,
    );

    if (pluginDestructureDirectRe.test(content) || pluginDestructureDirectBracketRe.test(content)) tauriPluginRoots.add("plugin");
    if (pluginsDestructureDirectRe.test(content) || pluginsDestructureDirectBracketRe.test(content)) tauriPluginsRoots.add("plugins");

    const pluginDestructureRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
      "g",
    );
    const pluginsDestructureRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
      "g",
    );
    const pluginDestructureRenameBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]`,
      "g",
    );
    const pluginsDestructureRenameBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]__TAURI__['\\\"]\\s*\\]`,
      "g",
    );
    for (const re of [
      { which: "plugin", re: pluginDestructureRenameRe },
      { which: "plugins", re: pluginsDestructureRenameRe },
      { which: "plugin", re: pluginDestructureRenameBracketRe },
      { which: "plugins", re: pluginsDestructureRenameBracketRe },
    ] as const) {
      let match: RegExpExecArray | null = null;
      while ((match = re.re.exec(content)) != null) {
        const name = match[1];
        if (name && re.which === "plugin") tauriPluginRoots.add(name);
        if (name && re.which === "plugins") tauriPluginsRoots.add(name);
        if (match[0].length === 0) re.re.lastIndex += 1;
      }
    }

    // Nested destructuring from the global alias:
    //   const { __TAURI__: { plugin } } = g;
    //   const { __TAURI__: { plugins: p } } = g;
    const pluginNestedDestructureFromGlobalAliasDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugin\\b(?!\\s*:)[\\s\\S]*?\\}\\s*=\\s*${g}`,
    );
    if (pluginNestedDestructureFromGlobalAliasDirectRe.test(content)) tauriPluginRoots.add("plugin");
    const pluginNestedDestructureFromGlobalAliasRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${g}`,
      "g",
    );
    let nestedMatch: RegExpExecArray | null = null;
    while ((nestedMatch = pluginNestedDestructureFromGlobalAliasRenameRe.exec(content)) != null) {
      const name = nestedMatch[1];
      if (name) tauriPluginRoots.add(name);
      if (nestedMatch[0].length === 0) pluginNestedDestructureFromGlobalAliasRenameRe.lastIndex += 1;
    }

    const pluginsNestedDestructureFromGlobalAliasDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugins\\b(?!\\s*:)[\\s\\S]*?\\}\\s*=\\s*${g}`,
    );
    if (pluginsNestedDestructureFromGlobalAliasDirectRe.test(content)) tauriPluginsRoots.add("plugins");
    const pluginsNestedDestructureFromGlobalAliasRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${g}`,
      "g",
    );
    while ((nestedMatch = pluginsNestedDestructureFromGlobalAliasRenameRe.exec(content)) != null) {
      const name = nestedMatch[1];
      if (name) tauriPluginsRoots.add(name);
      if (nestedMatch[0].length === 0) pluginsNestedDestructureFromGlobalAliasRenameRe.lastIndex += 1;
    }
  }

  return { tauriRoots, tauriPluginRoots, tauriPluginsRoots };
}

function buildBannedResForGlobalAlias(globalAlias: string): RegExp[] {
  const r = escapeRegExp(globalAlias);
  const g = buildAnyCastableRefSource(r);
  const namespaces = ["event", "window", "dialog"] as const;
  const containers = ["plugin", "plugins"] as const;

  /** @type {RegExp[]} */
  const out: RegExp[] = [];

  // Nested destructuring directly from the global alias itself:
  //   const { __TAURI__: { dialog } } = g;
  out.push(
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${g}`,
    ),
  );

  // Destructuring from g.__TAURI__ / g["__TAURI__"]:
  //   const { dialog } = g.__TAURI__;
  //   const { plugin: { dialog } } = g["__TAURI__"];
  out.push(
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\b`,
    ),
  );
  out.push(
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]`,
    ),
  );

  for (const ns of namespaces) {
    // g.__TAURI__.event/window/dialog
    out.push(new RegExp(`${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`));
    out.push(new RegExp(`${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${ns}['"]\\s*\\]`));
    // g["__TAURI__"].event/window/dialog
    out.push(new RegExp(`${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`));
    out.push(
      new RegExp(
        `${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${ns}['"]\\s*\\]`,
      ),
    );

    for (const container of containers) {
      // g.__TAURI__.plugin.event etc
      out.push(
        new RegExp(
          `${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.|\\.)\\s*${container}\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`,
        ),
      );
      out.push(
        new RegExp(
          `${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*${container}\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`,
        ),
      );
      // g.__TAURI__["plugin"].event etc
      out.push(
        new RegExp(
          `${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${container}['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`,
        ),
      );
      out.push(
        new RegExp(
          `${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${container}['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*${ns}\\b`,
        ),
      );

      // Destructuring directly from plugin containers:
      //   const { dialog } = g.__TAURI__.plugin;
      //   const { dialog } = g["__TAURI__"].plugin;
      out.push(
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b${ns}\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.|\\.)\\s*${container}\\b`,
        ),
      );
      out.push(
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b${ns}\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*${container}\\b`,
        ),
      );
      out.push(
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b${ns}\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.|\\.)\\s*__TAURI__\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${container}['"]\\s*\\]`,
        ),
      );
      out.push(
        new RegExp(
          `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b${ns}\\b[\\s\\S]*?\\}\\s*=\\s*${g}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]__TAURI__['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]${container}['"]\\s*\\]`,
        ),
      );
    }
  }

  return out;
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

  // Fast-path: most source files never mention the Tauri globals. Avoid running the heavier
  // regex scan in that case so this guard test stays cheap.
  if (!content.includes("__TAURI__")) {
    return { tauriRoots, tauriPluginRoots, tauriPluginsRoots };
  }

  // Capture common aliasing patterns like:
  //   const tauri = (globalThis as any).__TAURI__;
  //   let tauri = globalThis.__TAURI__ ?? null;
  //
  // NOTE: This intentionally only targets direct aliases to the root `__TAURI__` object (not
  // nested properties like `__TAURI__.core.invoke`), so we can then flag `tauri.dialog` /
  // `tauri.window` / `tauri.event` access even when the file doesn't mention `__TAURI__` again.
  const tauriRootAssignRe =
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__\b|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__\b|__TAURI__\b|\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]|(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\])(?!(?:\s*(?:\?\.|\.|\[)))/g;

  let match: RegExpExecArray | null = null;
  while ((match = tauriRootAssignRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriRoots.add(name);
    if (match[0].length === 0) tauriRootAssignRe.lastIndex += 1;
  }

  // Capture destructuring aliasing patterns like:
  //   const { __TAURI__: tauri } = globalThis;
  //   const { "__TAURI__": tauri } = (globalThis as any);
  const tauriRootDestructureRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}\\b`,
    "g",
  );
  const tauriRootDestructureRenameQuotedRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?['\\\"]__TAURI__['\\\"]\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}\\b`,
    "g",
  );
  for (const re of [tauriRootDestructureRenameRe, tauriRootDestructureRenameQuotedRe]) {
    while ((match = re.exec(content)) != null) {
      const name = match[1];
      if (name) tauriRoots.add(name);
      if (match[0].length === 0) re.lastIndex += 1;
    }
  }

  // Nested destructuring from the global object:
  //   const { __TAURI__: { plugin } } = globalThis;
  //   const { "__TAURI__": { plugins: p } } = (globalThis as any);
  const pluginNestedDestructureFromGlobalDirectRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugin\\b(?!\\s*:)[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}`,
  );
  if (pluginNestedDestructureFromGlobalDirectRe.test(content)) tauriPluginRoots.add("plugin");
  const pluginNestedDestructureFromGlobalRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}`,
    "g",
  );
  while ((match = pluginNestedDestructureFromGlobalRenameRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriPluginRoots.add(name);
    if (match[0].length === 0) pluginNestedDestructureFromGlobalRenameRe.lastIndex += 1;
  }

  const pluginsNestedDestructureFromGlobalDirectRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugins\\b(?!\\s*:)[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}`,
  );
  if (pluginsNestedDestructureFromGlobalDirectRe.test(content)) tauriPluginsRoots.add("plugins");
  const pluginsNestedDestructureFromGlobalRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}`,
    "g",
  );
  while ((match = pluginsNestedDestructureFromGlobalRenameRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriPluginsRoots.add(name);
    if (match[0].length === 0) pluginsNestedDestructureFromGlobalRenameRe.lastIndex += 1;
  }

  // Capture aliases to the plugin container objects:
  //   const plugin = (globalThis as any).__TAURI__?.plugin;
  const tauriPluginAssignRe =
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__|__TAURI__|\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]|(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\])\s*(?:\?\.|\.)\s*(plugin|plugins)\b(?!\s*(?:\?\.|\.|\[))/g;

  while ((match = tauriPluginAssignRe.exec(content)) != null) {
    const name = match[1];
    const which = match[2];
    if (name && which === "plugin") tauriPluginRoots.add(name);
    if (name && which === "plugins") tauriPluginsRoots.add(name);
    if (match[0].length === 0) tauriPluginAssignRe.lastIndex += 1;
  }

  // Capture plugin container aliases via destructuring:
  //   const { plugin } = __TAURI__;
  //   const { plugin: p } = tauri;
  const pluginDestructureDirectRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b(?!\\s*:)` + `[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}`,
  );
  if (pluginDestructureDirectRe.test(content)) tauriPluginRoots.add("plugin");
  const pluginDestructureRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}`,
    "g",
  );
  while ((match = pluginDestructureRenameRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriPluginRoots.add(name);
    if (match[0].length === 0) pluginDestructureRenameRe.lastIndex += 1;
  }

  const pluginsDestructureDirectRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b(?!\\s*:)` + `[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}`,
  );
  if (pluginsDestructureDirectRe.test(content)) tauriPluginsRoots.add("plugins");
  const pluginsDestructureRenameRe = new RegExp(
    `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${TAURI_GLOBAL_REF_RE_SOURCE}`,
    "g",
  );
  while ((match = pluginsDestructureRenameRe.exec(content)) != null) {
    const name = match[1];
    if (name) tauriPluginsRoots.add(name);
    if (match[0].length === 0) pluginsDestructureRenameRe.lastIndex += 1;
  }

  for (const root of tauriRoots) {
    const r = escapeRegExp(root);

    const pluginAssignFromAliasDotRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\b`,
      "g",
    );
    const pluginsAssignFromAliasDotRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\b`,
      "g",
    );
    const pluginAssignFromAliasBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]plugin['\\\"]\\s*\\]`,
      "g",
    );
    const pluginsAssignFromAliasBracketRe = new RegExp(
      `\\b(?:const|let|var)\\s+([A-Za-z_$][\\w$]*)\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['\\\"]plugins['\\\"]\\s*\\]`,
      "g",
    );

    for (const re of [
      { which: "plugin", re: pluginAssignFromAliasDotRe },
      { which: "plugins", re: pluginsAssignFromAliasDotRe },
      { which: "plugin", re: pluginAssignFromAliasBracketRe },
      { which: "plugins", re: pluginsAssignFromAliasBracketRe },
    ] as const) {
      while ((match = re.re.exec(content)) != null) {
        const name = match[1];
        if (!name) continue;
        if (re.which === "plugin") tauriPluginRoots.add(name);
        else tauriPluginsRoots.add(name);
        if (match[0].length === 0) re.re.lastIndex += 1;
      }
    }

    // Destructuring from a root alias: `const { plugin } = tauri;`
    const pluginDestructureFromAliasDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b(?!\\s*:)[^}]*\\}\\s*=\\s*${r}\\b`,
    );
    if (pluginDestructureFromAliasDirectRe.test(content)) tauriPluginRoots.add("plugin");
    const pluginDestructureFromAliasRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugin\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${r}\\b`,
      "g",
    );
    while ((match = pluginDestructureFromAliasRenameRe.exec(content)) != null) {
      const name = match[1];
      if (name) tauriPluginRoots.add(name);
      if (match[0].length === 0) pluginDestructureFromAliasRenameRe.lastIndex += 1;
    }

    const pluginsDestructureFromAliasDirectRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b(?!\\s*:)[^}]*\\}\\s*=\\s*${r}\\b`,
    );
    if (pluginsDestructureFromAliasDirectRe.test(content)) tauriPluginsRoots.add("plugins");
    const pluginsDestructureFromAliasRenameRe = new RegExp(
      `\\b(?:const|let|var)\\s*\\{[^}]*\\bplugins\\b\\s*:\\s*([A-Za-z_$][\\w$]*)\\b[^}]*\\}\\s*=\\s*${r}\\b`,
      "g",
    );
    while ((match = pluginsDestructureFromAliasRenameRe.exec(content)) != null) {
      const name = match[1];
      if (name) tauriPluginsRoots.add(name);
      if (match[0].length === 0) pluginsDestructureFromAliasRenameRe.lastIndex += 1;
    }
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
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`),
    // Destructuring patterns: `const { dialog } = tauri;`
    new RegExp(`\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\b`),

    // Plugin container variants.
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`),
    new RegExp(
      `\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`,
    ),
    new RegExp(
      `\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`,
    ),
    new RegExp(
      `\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`,
    ),

    // Destructuring directly from plugin containers (not via an intermediate alias):
    //   const { dialog } = tauri.plugin;
    //   const { dialog } = tauri["plugin"];
    //   const { dialog } = tauri.plugins;
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*plugin\\b`,
    ),
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugin['"]\\s*\\]`,
    ),
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.|\\.)\\s*plugins\\b`,
    ),
    new RegExp(
      `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]plugins['"]\\s*\\]`,
    ),
  ];
}

function buildBannedResForTauriPluginAlias(root: string): RegExp[] {
  const r = escapeRegExp(root);
  return [
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*event\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*window\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.|\\.)\\s*dialog\\b`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]event['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]window['"]\\s*\\]`),
    new RegExp(`\\b${r}\\s*(?:\\?\\.)?\\s*\\[\\s*['"]dialog['"]\\s*\\]`),
    new RegExp(`\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${r}\\b`),
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
      // Bracket access to the __TAURI__ global itself (e.g. globalThis["__TAURI__"].event).
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*event\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*event\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*event\b/,
      // Bracket access variants: __TAURI__["event"] / __TAURI__?.["event"] / mixed plugin container shapes.
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*event\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*event\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*event\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*event\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]event['"]\s*\]/,

      // Window API access should go through getTauriWindowHandleOr{Null,Throw} or hasTauriWindow* helpers.
      /\b__TAURI__\s*(?:\?\.)\s*window\b/,
      /\b__TAURI__\s*\.\s*window\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*window\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*window\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*window\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*window\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*window\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*window\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]window['"]\s*\]/,

      // Dialog API access should go through tauri/api helpers (or `nativeDialogs` where appropriate).
      /\b__TAURI__\s*(?:\?\.)\s*dialog\b/,
      /\b__TAURI__\s*\.\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*dialog\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*dialog\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.|\.)\s*dialog\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.|\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*dialog\b/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b__TAURI__\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.|\.)\s*dialog\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.|\.)\s*dialog\b/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugin\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.|\.)\s*plugins\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugin['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,
      /\b(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]plugins['"]\s*\]\s*(?:\?\.)?\s*\[\s*['"]dialog['"]\s*\]/,

      // Destructuring patterns: `const { dialog } = globalThis.__TAURI__;`
      /\b(?:const|let|var)\s*\{[\s\S]*?\b(?:dialog|event|window)\b[\s\S]*?\}\s*=\s*(?:\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.|\.)\s*__TAURI__\b|(?:globalThis|window|self)\s*(?:\?\.|\.)\s*__TAURI__\b|__TAURI__\b|\(\s*(?:globalThis|window|self)\s+as\s+any\s*\)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\]|(?:globalThis|window|self)\s*(?:\?\.)?\s*\[\s*['"]__TAURI__['"]\s*\])/,

      // Nested destructuring from the global object:
      //   const { __TAURI__: { dialog } } = globalThis;
      //   const { "__TAURI__": { plugin: { window } } } = (window as any);
      new RegExp(
        `\\b(?:const|let|var)\\s*\\{[\\s\\S]*?\\b__TAURI__\\b[\\s\\S]*?\\b(?:dialog|event|window)\\b[\\s\\S]*?\\}\\s*=\\s*${GLOBAL_OBJECT_REF_RE_SOURCE}`,
      ),
    ];

    for (const absPath of files) {
      const relPath = path.relative(SRC_ROOT, absPath);
      if (isTestFile(relPath)) continue;

      const normalized = relPath.replace(/\\/g, "/");
      if (normalized === "tauri/api.ts" || normalized === "tauri/api.js") continue;

      // Strip comments so commented-out `__TAURI__` API access cannot satisfy or fail this guardrail.
      const raw = await readFile(absPath, "utf8");
      const content = stripComments(raw);
      // Fast-path: if the file doesn't mention the Tauri globals at all, none of the banned
      // patterns can match (including the alias-based checks in this guard).
      if (!content.includes("__TAURI__")) continue;
      const rawLines = raw.split(/\r?\n/);

      const globalAliases = collectGlobalObjectAliases(content);
      const dynamicBannedRes: RegExp[] = [];
      for (const alias of globalAliases) {
        dynamicBannedRes.push(...buildBannedResForGlobalAlias(alias));
      }

      const aliasRes: RegExp[] = [];
      const aliases = collectTauriAliases(content);
      const extraAliases = collectTauriAliasesFromGlobalAliases(content, globalAliases);
      for (const alias of extraAliases.tauriRoots) aliases.tauriRoots.add(alias);
      for (const alias of extraAliases.tauriPluginRoots) aliases.tauriPluginRoots.add(alias);
      for (const alias of extraAliases.tauriPluginsRoots) aliases.tauriPluginsRoots.add(alias);
      for (const alias of aliases.tauriRoots) {
        aliasRes.push(...buildBannedResForTauriAlias(alias));
      }
      for (const alias of aliases.tauriPluginRoots) {
        aliasRes.push(...buildBannedResForTauriPluginAlias(alias));
      }
      for (const alias of aliases.tauriPluginsRoots) {
        aliasRes.push(...buildBannedResForTauriPluginAlias(alias));
      }

      for (const re of [...bannedRes, ...dynamicBannedRes, ...aliasRes]) {
        const globalRe = new RegExp(re.source, re.flags.includes("g") ? re.flags : `${re.flags}g`);
        let match: RegExpExecArray | null = null;
        while ((match = globalRe.exec(content)) != null) {
          const start = match.index;
          const lineNumber = content.slice(0, start).split(/\r?\n/).length;
          const line = rawLines[lineNumber - 1] ?? "";
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
