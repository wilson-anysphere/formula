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

function normalizeQuery(query: string): string {
  return String(query ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function normalizeCategory(category: string | null): string {
  const value = String(category ?? "").trim();
  return value ? value : "Other";
}

function getPrimaryShortcut(value: string | readonly string[] | undefined): string | null {
  if (!value) return null;
  if (typeof value === "string") return value;
  if (Array.isArray(value)) return value[0] ?? null;
  return null;
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
}): Array<ShortcutSearchMatch<T>> {
  const q = normalizeQuery(params.query);

  const matches: Array<ShortcutSearchMatch<T>> = [];
  for (const cmd of params.commands) {
    const shortcut = getPrimaryShortcut(params.keybindingIndex.get(cmd.commandId));
    if (!shortcut) continue;

    if (q) {
      const haystack = `${cmd.title} ${cmd.commandId} ${shortcut}`.toLowerCase();
      if (!haystack.includes(q)) continue;
    }

    matches.push({ ...cmd, shortcut });
  }

  matches.sort((a, b) => {
    const catA = normalizeCategory(a.category);
    const catB = normalizeCategory(b.category);
    const catCompare = catA.localeCompare(catB, undefined, { sensitivity: "base" });
    if (catCompare !== 0) return catCompare;

    const shortcutCompare = a.shortcut.localeCompare(b.shortcut, undefined, { sensitivity: "base" });
    if (shortcutCompare !== 0) return shortcutCompare;

    const titleCompare = a.title.localeCompare(b.title, undefined, { sensitivity: "base" });
    if (titleCompare !== 0) return titleCompare;

    return a.commandId.localeCompare(b.commandId, undefined, { sensitivity: "base" });
  });

  return matches;
}
