export const COMMAND_PALETTE_RECENTS_STORAGE_KEY = "formula.commandPalette.recents";

export type StorageLike = Pick<Storage, "getItem" | "setItem">;

export function readCommandPaletteRecents(
  storage: StorageLike,
  { storageKey = COMMAND_PALETTE_RECENTS_STORAGE_KEY }: { storageKey?: string } = {},
): string[] {
  try {
    const raw = storage.getItem(storageKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];

    const out: string[] = [];
    const seen = new Set<string>();
    for (const item of parsed) {
      if (typeof item !== "string") continue;
      const id = item.trim();
      if (!id) continue;
      if (seen.has(id)) continue;
      seen.add(id);
      out.push(id);
    }
    return out;
  } catch {
    return [];
  }
}

export function writeCommandPaletteRecents(
  storage: StorageLike,
  ids: string[],
  { storageKey = COMMAND_PALETTE_RECENTS_STORAGE_KEY }: { storageKey?: string } = {},
): void {
  try {
    storage.setItem(storageKey, JSON.stringify(ids));
  } catch {
    // Ignore storage failures (private mode, quota, etc.)
  }
}

export function recordCommandPaletteRecent(
  storage: StorageLike,
  commandId: string,
  {
    limit = 25,
    storageKey = COMMAND_PALETTE_RECENTS_STORAGE_KEY,
  }: { limit?: number; storageKey?: string } = {},
): string[] {
  const id = String(commandId ?? "").trim();
  if (!id) return readCommandPaletteRecents(storage, { storageKey });

  const current = readCommandPaletteRecents(storage, { storageKey });
  const next = [id, ...current.filter((c) => c !== id)].slice(0, Math.max(0, limit));
  writeCommandPaletteRecents(storage, next, { storageKey });
  return next;
}

