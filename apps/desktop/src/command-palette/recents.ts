import type { CommandRegistry } from "../extensions/commandRegistry.js";

export const COMMAND_RECENTS_STORAGE_KEY = "formula.commandRecents";
export const LEGACY_COMMAND_RECENTS_STORAGE_KEY = "formula.commandPalette.recents";
export const DEFAULT_COMMAND_RECENTS_MAX_ENTRIES = 20;

export type StorageLike = Pick<Storage, "getItem" | "setItem">;

export type CommandRecentEntry = {
  commandId: string;
  lastUsedMs: number;
  count?: number;
};

function safeParseRecents(raw: string | null): CommandRecentEntry[] {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];

    const out: Array<CommandRecentEntry & { __order: number }> = [];
    const seen = new Set<string>();
    let order = 0;
    for (const item of parsed) {
      if (!item || typeof item !== "object") continue;
      const commandId = typeof (item as any).commandId === "string" ? String((item as any).commandId).trim() : "";
      const lastUsedMs = typeof (item as any).lastUsedMs === "number" ? (item as any).lastUsedMs : NaN;
      const count =
        typeof (item as any).count === "number" && Number.isFinite((item as any).count) ? (item as any).count : undefined;
      if (!commandId) continue;
      if (!Number.isFinite(lastUsedMs)) continue;
      if (seen.has(commandId)) continue;
      seen.add(commandId);
      out.push({ commandId, lastUsedMs, ...(count != null ? { count } : {}), __order: order++ });
    }

    // Ensure deterministic ordering for callers (and for the stored JSON).
    out.sort((a, b) => b.lastUsedMs - a.lastUsedMs || a.__order - b.__order);
    return out.map(({ __order: _order, ...entry }) => entry);
  } catch {
    return [];
  }
}

function safeParseLegacyRecents(raw: string | null): string[] {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];

    const out: string[] = [];
    const seen = new Set<string>();
    for (const item of parsed) {
      const commandId =
        typeof item === "string"
          ? item
          : item && typeof item === "object" && typeof (item as any).commandId === "string"
            ? String((item as any).commandId)
            : "";
      const normalized = String(commandId).trim();
      if (!normalized) continue;
      if (seen.has(normalized)) continue;
      seen.add(normalized);
      out.push(normalized);
    }
    return out;
  } catch {
    return [];
  }
}

export function readCommandRecents(
  storage: StorageLike,
  { storageKey = COMMAND_RECENTS_STORAGE_KEY }: { storageKey?: string } = {},
): CommandRecentEntry[] {
  try {
    return safeParseRecents(storage.getItem(storageKey));
  } catch {
    return [];
  }
}

export function writeCommandRecents(
  storage: StorageLike,
  entries: CommandRecentEntry[],
  { storageKey = COMMAND_RECENTS_STORAGE_KEY }: { storageKey?: string } = {},
): void {
  try {
    storage.setItem(storageKey, JSON.stringify(entries));
  } catch {
    // Ignore storage failures (private mode, quota, etc.)
  }
}

export function recordCommandRecent(
  storage: StorageLike,
  commandId: string,
  {
    maxEntries = DEFAULT_COMMAND_RECENTS_MAX_ENTRIES,
    nowMs = Date.now(),
    storageKey = COMMAND_RECENTS_STORAGE_KEY,
  }: { maxEntries?: number; nowMs?: number; storageKey?: string } = {},
): CommandRecentEntry[] {
  const id = String(commandId ?? "").trim();
  if (!id) return readCommandRecents(storage, { storageKey });

  const limit = Number.isFinite(maxEntries) ? Math.max(0, maxEntries) : DEFAULT_COMMAND_RECENTS_MAX_ENTRIES;
  const now = typeof nowMs === "number" && Number.isFinite(nowMs) ? nowMs : Date.now();

  const current = readCommandRecents(storage, { storageKey });
  const without = current.filter((entry) => entry.commandId !== id);
  const prev = current.find((entry) => entry.commandId === id) ?? null;

  const next: Array<CommandRecentEntry & { __order: number }> = [
    { commandId: id, lastUsedMs: now, count: (prev?.count ?? 0) + 1, __order: 0 },
    ...without.map((entry, idx) => ({ ...entry, __order: idx + 1 })),
  ];
  next.sort((a, b) => b.lastUsedMs - a.lastUsedMs || a.__order - b.__order);

  const trimmed =
    limit > 0 ? next.slice(0, limit).map(({ __order: _order, ...entry }) => entry) : [];
  writeCommandRecents(storage, trimmed, { storageKey });
  return trimmed;
}

export function getRecentCommandIdsForDisplay(
  storage: StorageLike,
  existingCommandIds: Iterable<string>,
  {
    limit = 8,
    storageKey = COMMAND_RECENTS_STORAGE_KEY,
  }: { limit?: number; storageKey?: string } = {},
): string[] {
  const ids = new Set(Array.from(existingCommandIds, (id) => String(id)));
  if (ids.size === 0) return [];

  const recents = readCommandRecents(storage, { storageKey });
  const out: string[] = [];
  const max = Number.isFinite(limit) ? Math.max(0, limit) : 0;
  if (max === 0) return out;

  for (const entry of recents) {
    if (!ids.has(entry.commandId)) continue;
    out.push(entry.commandId);
    if (out.length >= max) break;
  }
  return out;
}

export function installCommandRecentsTracker(
  commandRegistry: Pick<CommandRegistry, "onDidExecuteCommand">,
  storage: StorageLike,
  options: { maxEntries?: number; now?: () => number; storageKey?: string; ignoreCommandIds?: readonly string[] } = {},
): () => void {
  const now = options.now ?? (() => Date.now());
  const ignore = new Set((options.ignoreCommandIds ?? []).map((id) => String(id).trim()).filter(Boolean));
  const storageKey = options.storageKey ?? COMMAND_RECENTS_STORAGE_KEY;
  const maxEntries = Number.isFinite(options.maxEntries)
    ? Math.max(0, Math.floor(options.maxEntries ?? DEFAULT_COMMAND_RECENTS_MAX_ENTRIES))
    : DEFAULT_COMMAND_RECENTS_MAX_ENTRIES;

  // Some commands are registered purely for ribbon/schema compatibility and are intentionally
  // hidden from the command palette via `when: "false"`. Do not record them as recents: they
  // would never be shown in the palette and can crowd out meaningful entries in storage.
  const maybeGetCommand = (() => {
    const registry: any = commandRegistry;
    if (!registry || typeof registry.getCommand !== "function") return null;
    return (commandId: string): { when?: string | null } | null => {
      try {
        return registry.getCommand(commandId) ?? null;
      } catch {
        return null;
      }
    };
  })();
  const isAlwaysHiddenFromPalette = (commandId: string): boolean => {
    if (!maybeGetCommand) return false;
    const when = maybeGetCommand(commandId)?.when;
    // `evaluateWhenClause` treats boolean literals case-insensitively, so match that here
    // to ensure we ignore hidden commands even if authors accidentally register `when: "False"`.
    return typeof when === "string" && when.trim().toLowerCase() === "false";
  };

  // Best-effort, one-time migration from the legacy recents key.
  // We only migrate when the new key has no entries yet, so it is idempotent.
  try {
    let existing = readCommandRecents(storage, { storageKey });

    // If the ignore list changes over time (e.g. we start ignoring clipboard commands),
    // drop ignored entries eagerly so the "RECENT" group stays useful immediately after update.
    if (existing.length > 0) {
      const filtered = existing.filter((entry) => !ignore.has(entry.commandId) && !isAlwaysHiddenFromPalette(entry.commandId));
      if (filtered.length !== existing.length) {
        writeCommandRecents(storage, filtered, { storageKey });
        existing = filtered;
      }
    }

    // Enforce the configured size cap even if storage already contains more entries.
    if (existing.length > maxEntries) {
      const trimmed = maxEntries > 0 ? existing.slice(0, maxEntries) : [];
      writeCommandRecents(storage, trimmed, { storageKey });
      existing = trimmed;
    }

    if (existing.length === 0) {
      const legacyIds = safeParseLegacyRecents(storage.getItem(LEGACY_COMMAND_RECENTS_STORAGE_KEY)).filter(
        (id) => !ignore.has(id) && !isAlwaysHiddenFromPalette(id),
      );
      if (legacyIds.length > 0) {
        const nowMs = now();
        const migrated: CommandRecentEntry[] = (maxEntries > 0 ? legacyIds.slice(0, maxEntries) : []).map(
          (commandId) => ({
            commandId,
            lastUsedMs: nowMs,
            count: 1,
          }),
        );
        if (migrated.length > 0) writeCommandRecents(storage, migrated, { storageKey });
      }
    }
  } catch {
    // ignore
  }

  return commandRegistry.onDidExecuteCommand((evt) => {
    const commandId = String(evt.commandId ?? "").trim();
    if (!commandId) return;
    if (ignore.has(commandId)) return;
    if (isAlwaysHiddenFromPalette(commandId)) return;
    // Only record successful executions. `CommandRegistry` emits either `{ result }` or `{ error }`.
    // Use result presence, not `evt.error == null`, so `throw undefined` doesn't count as success.
    if (!("result" in evt)) return;
    recordCommandRecent(storage, commandId, {
      maxEntries,
      nowMs: now(),
      storageKey,
    });
  });
}
