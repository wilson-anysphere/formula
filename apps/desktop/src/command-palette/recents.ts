import type { CommandRegistry } from "../extensions/commandRegistry.js";

export const COMMAND_RECENTS_STORAGE_KEY = "formula.commandRecents";
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

    const out: CommandRecentEntry[] = [];
    const seen = new Set<string>();
    for (const item of parsed) {
      if (!item || typeof item !== "object") continue;
      const commandId = typeof (item as any).commandId === "string" ? String((item as any).commandId).trim() : "";
      const lastUsedMs = typeof (item as any).lastUsedMs === "number" ? (item as any).lastUsedMs : NaN;
      const count = typeof (item as any).count === "number" ? (item as any).count : undefined;
      if (!commandId) continue;
      if (!Number.isFinite(lastUsedMs)) continue;
      if (seen.has(commandId)) continue;
      seen.add(commandId);
      out.push({ commandId, lastUsedMs, ...(count != null ? { count } : {}) });
    }

    // Ensure deterministic ordering for callers (and for the stored JSON).
    out.sort((a, b) => b.lastUsedMs - a.lastUsedMs);
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

  const next: CommandRecentEntry[] = [
    { commandId: id, lastUsedMs: now, count: (prev?.count ?? 0) + 1 },
    ...without,
  ];
  next.sort((a, b) => b.lastUsedMs - a.lastUsedMs);

  const trimmed = limit > 0 ? next.slice(0, limit) : [];
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
  const ignore = new Set((options.ignoreCommandIds ?? []).map((id) => String(id)));

  return commandRegistry.onDidExecuteCommand((evt) => {
    if (ignore.has(evt.commandId)) return;
    recordCommandRecent(storage, evt.commandId, {
      maxEntries: options.maxEntries,
      nowMs: now(),
      storageKey: options.storageKey,
    });
  });
}
