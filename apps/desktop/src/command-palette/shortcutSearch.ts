export type ShortcutSearchCommandLike = {
  commandId: string;
  title: string;
  category: string | null;
};

export type ShortcutSearchMatch<T extends ShortcutSearchCommandLike = ShortcutSearchCommandLike> = T & {
  /**
   * Primary (display) shortcut string for the command.
   */
  shortcut: string;
};

const SHORTCUT_SYMBOL_TOKENS: Record<string, string> = {
  // Modifiers.
  "⌘": "cmd",
  "⇧": "shift",
  "⌥": "alt",
  "⌃": "ctrl",

  // Common keys.
  "⎋": "escape",
  "↩": "enter",
  "↵": "enter",
  "⌫": "backspace",
  "⌦": "delete",
  "⇥": "tab",

  // Arrows.
  "↑": "up",
  "↓": "down",
  "←": "left",
  "→": "right",
};

// Keys that appear as literal single-character tokens (e.g. "[" in "⌘[").
// We intentionally include common punctuation used by spreadsheet shortcuts.
const SHORTCUT_PUNCTUATION_KEYS = new Set([
  "[",
  "]",
  ";",
  "'",
  ",",
  ".",
  "/",
  "\\",
  "`",
  "-",
  "=",
  "$",
  "%",
  "#",
]);

type ShortcutSearchLimits = {
  /**
   * Max number of matches to return.
   * When provided (and query is empty), we avoid sorting potentially huge match arrays.
   */
  maxResults: number;
  /**
   * Max number of matches per category to return.
   */
  maxResultsPerCategory: number;
};

const SHORTCUT_TOKEN_SYNONYMS: Record<string, string> = {
  // Modifiers.
  cmd: "cmd",
  command: "cmd",
  meta: "cmd",
  win: "cmd",
  super: "cmd",

  shift: "shift",

  ctrl: "ctrl",
  control: "ctrl",

  alt: "alt",
  option: "alt",
  opt: "alt",

  // Common keys.
  esc: "escape",
  escape: "escape",
  return: "enter",
  enter: "enter",
  tab: "tab",
  space: "space",
  spacebar: "space",
  del: "delete",
  delete: "delete",
  backspace: "backspace",

  // Arrows.
  up: "up",
  arrowup: "up",
  down: "down",
  arrowdown: "down",
  left: "left",
  arrowleft: "left",
  right: "right",
  arrowright: "right",

  // Paging keys.
  pageup: "pageup",
  pgup: "pageup",
  pagedown: "pagedown",
  pgdn: "pagedown",
};

import { t } from "../i18n/index.js";

function normalizeQuery(query: string): string {
  return String(query ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function normalizeCategory(category: string | null): string {
  const value = String(category ?? "").trim();
  return value ? value : t("commandPalette.group.other");
}

function extractShortcutTokens(text: string): string[] {
  const tokens: string[] = [];
  let buffer = "";

  const normalizeToken = (token: string): string => {
    const lower = token.toLowerCase();
    return SHORTCUT_TOKEN_SYNONYMS[lower] ?? lower;
  };

  const flush = () => {
    if (!buffer) return;
    tokens.push(normalizeToken(buffer));
    buffer = "";
  };

  for (const ch of String(text ?? "")) {
    const mapped = SHORTCUT_SYMBOL_TOKENS[ch];
    if (mapped) {
      flush();
      tokens.push(normalizeToken(mapped));
      continue;
    }

    if (/[a-z0-9]/i.test(ch)) {
      buffer += ch;
      continue;
    }

    flush();

    if (SHORTCUT_PUNCTUATION_KEYS.has(ch)) {
      tokens.push(normalizeToken(ch));
    }
  }

  flush();
  return tokens;
}

function matchesShortcutTokenQuery(shortcutDisplay: string, query: string): boolean {
  const qTokens = extractShortcutTokens(query);
  if (qTokens.length === 0) return false;

  const shortcutTokens = new Set(extractShortcutTokens(shortcutDisplay));
  for (const token of qTokens) {
    if (!shortcutTokens.has(token)) return false;
  }
  return true;
}

function getPrimaryShortcut(value: string | readonly string[] | undefined): string | null {
  if (!value) return null;
  if (typeof value === "string") return value;
  if (Array.isArray(value)) return value[0] ?? null;
  return null;
}

function compareShortcutMatches(a: { shortcut: string; title: string; commandId: string }, b: { shortcut: string; title: string; commandId: string }): number {
  const shortcutCompare = a.shortcut.localeCompare(b.shortcut, undefined, { sensitivity: "base" });
  if (shortcutCompare !== 0) return shortcutCompare;

  const titleCompare = a.title.localeCompare(b.title, undefined, { sensitivity: "base" });
  if (titleCompare !== 0) return titleCompare;

  return a.commandId.localeCompare(b.commandId, undefined, { sensitivity: "base" });
}

/**
 * Returns the commands eligible for shortcut search:
 * - must have a display shortcut in `keybindingIndex`
 * - optional filtering by query (title / id / shortcut display)
 * - sorted by category, then shortcut string, then title
 */
export function searchShortcutCommands<T extends ShortcutSearchCommandLike>(params: {
  commands: T[];
  keybindingIndex: Map<string, string | readonly string[]>;
  query: string;
  limits?: ShortcutSearchLimits;
}): Array<ShortcutSearchMatch<T>> {
  const q = normalizeQuery(params.query);

  const limits = params.limits ?? null;
  const maxResults = limits ? Math.max(0, Math.floor(limits.maxResults)) : null;
  const maxPerCategoryRaw = limits ? Math.max(1, Math.floor(limits.maxResultsPerCategory)) : null;
  const maxPerCategory = maxResults != null && maxPerCategoryRaw != null ? Math.min(maxResults, maxPerCategoryRaw) : null;

  // When limits are provided we can avoid sorting potentially huge match arrays by
  // keeping only the best N shortcuts per category and returning at most maxResults.
  if (maxResults != null && maxPerCategory != null) {
    if (maxResults === 0) return [];

    const byCategory = new Map<string, Array<ShortcutSearchMatch<T>>>();
    for (const cmd of params.commands) {
      const shortcut = getPrimaryShortcut(params.keybindingIndex.get(cmd.commandId));
      if (!shortcut) continue;

      if (q) {
        const haystack = `${cmd.title} ${cmd.commandId} ${shortcut}`.toLowerCase();
        if (!haystack.includes(q) && !matchesShortcutTokenQuery(shortcut, q)) continue;
      }

      const category = normalizeCategory(cmd.category);
      const list = byCategory.get(category) ?? [];

      const candidate: ShortcutSearchMatch<T> = { ...cmd, shortcut };
      if (list.length < maxPerCategory) {
        list.push(candidate);
        byCategory.set(category, list);
        continue;
      }

      // Replace the worst entry in this category if the new one sorts earlier.
      let worstIdx = 0;
      for (let i = 1; i < list.length; i += 1) {
        if (compareShortcutMatches(list[i]!, list[worstIdx]!) > 0) worstIdx = i;
      }

      if (compareShortcutMatches(candidate, list[worstIdx]!) < 0) {
        list[worstIdx] = candidate;
        byCategory.set(category, list);
      }
    }

    const categories = [...byCategory.keys()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    const out: Array<ShortcutSearchMatch<T>> = [];
    for (const category of categories) {
      if (out.length >= maxResults) break;
      const list = byCategory.get(category);
      if (!list || list.length === 0) continue;
      list.sort(compareShortcutMatches);
      for (const cmd of list) {
        out.push(cmd);
        if (out.length >= maxResults) break;
      }
    }
    return out;
  }

  const matches: Array<ShortcutSearchMatch<T>> = [];
  for (const cmd of params.commands) {
    const shortcut = getPrimaryShortcut(params.keybindingIndex.get(cmd.commandId));
    if (!shortcut) continue;

    if (q) {
      const haystack = `${cmd.title} ${cmd.commandId} ${shortcut}`.toLowerCase();
      if (!haystack.includes(q) && !matchesShortcutTokenQuery(shortcut, q)) continue;
    }

    matches.push({ ...cmd, shortcut });
  }

  matches.sort((a, b) => {
    const catA = normalizeCategory(a.category);
    const catB = normalizeCategory(b.category);
    const catCompare = catA.localeCompare(catB, undefined, { sensitivity: "base" });
    if (catCompare !== 0) return catCompare;
    return compareShortcutMatches(a, b);
  });

  return matches;
}
