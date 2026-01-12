export type DockSide = "left" | "right" | "bottom";

export type ContributedPanelSeed = {
  extensionId: string;
  title: string;
  icon?: string | null;
  defaultDock?: DockSide;
};

export type ContributedPanelsSeedStoreData = Record<string, ContributedPanelSeed>;

// Versioned key so we can migrate safely later without breaking persisted layouts.
export const CONTRIBUTED_PANELS_SEED_STORE_KEY = "formula.extensions.contributedPanels.v1";

type StorageLike = Pick<Storage, "getItem" | "setItem" | "removeItem">;

function safeGetItem(storage: StorageLike, key: string): string | null {
  try {
    return storage.getItem(key);
  } catch {
    return null;
  }
}

function safeSetItem(storage: StorageLike, key: string, value: string): void {
  storage.setItem(key, value);
}

function safeRemoveItem(storage: StorageLike, key: string): void {
  try {
    storage.removeItem(key);
  } catch {
    // ignore
  }
}

function normalizeDockSide(value: unknown): DockSide | undefined {
  if (value === "left" || value === "right" || value === "bottom") return value;
  return undefined;
}

function normalizeIcon(value: unknown): string | null | undefined {
  if (value === undefined) return undefined;
  if (value === null) return null;
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : null;
  }
  return undefined;
}

export function readContributedPanelsSeedStore(storage: StorageLike): ContributedPanelsSeedStoreData {
  const raw = safeGetItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
  if (raw == null) return {};

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    safeRemoveItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
    return {};
  }

  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    safeRemoveItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
    return {};
  }

  const out: ContributedPanelsSeedStoreData = {};
  for (const [panelId, value] of Object.entries(parsed as Record<string, unknown>)) {
    if (typeof panelId !== "string" || panelId.trim().length === 0) continue;
    if (!value || typeof value !== "object" || Array.isArray(value)) continue;
    const seed = value as Record<string, unknown>;
    const extensionId = typeof seed.extensionId === "string" ? seed.extensionId.trim() : "";
    if (!extensionId) continue;
    const title = typeof seed.title === "string" ? seed.title.trim() : "";
    if (!title) continue;
    const icon = normalizeIcon(seed.icon);
    const defaultDock = normalizeDockSide(seed.defaultDock);
    out[panelId] = {
      extensionId,
      title,
      ...(icon !== undefined ? { icon } : {}),
      ...(defaultDock ? { defaultDock } : {}),
    };
  }

  return out;
}

export function writeContributedPanelsSeedStore(storage: StorageLike, data: ContributedPanelsSeedStoreData): void {
  const normalized: ContributedPanelsSeedStoreData = {};
  for (const [panelId, seed] of Object.entries(data ?? {})) {
    if (typeof panelId !== "string" || panelId.trim().length === 0) continue;
    if (!seed || typeof seed !== "object") continue;
    const extensionId = typeof seed.extensionId === "string" ? seed.extensionId.trim() : "";
    const title = typeof seed.title === "string" ? seed.title.trim() : "";
    if (!extensionId || !title) continue;
    const icon = normalizeIcon((seed as any).icon);
    const defaultDock = normalizeDockSide((seed as any).defaultDock);
    normalized[panelId] = {
      extensionId,
      title,
      ...(icon !== undefined ? { icon } : {}),
      ...(defaultDock ? { defaultDock } : {}),
    };
  }
  // When empty, remove the key entirely so uninstall behaves like a clean slate (consistent with
  // WebExtensionManager's uninstall cleanup behavior).
  if (Object.keys(normalized).length === 0) {
    safeRemoveItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
    return;
  }
  safeSetItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY, JSON.stringify(normalized));
}

/**
 * Replace the contributed panel metadata for an extension.
 *
 * Used for install/update flows.
 *
 * Correctness:
 * - Existing entries for the extension are replaced atomically.
 * - Conflicts (two extensions claiming the same panel id) are rejected without mutating the store.
 */
export function setSeedPanelsForExtension(
  storage: StorageLike,
  extensionId: string,
  panels: Array<{ id?: unknown; title?: unknown; icon?: unknown; defaultDock?: unknown; position?: unknown }>,
  options: { onError?: (message: string) => void } = {},
): boolean {
  const owner = String(extensionId ?? "").trim();
  if (!owner) return false;

  const current = readContributedPanelsSeedStore(storage);
  const next: ContributedPanelsSeedStoreData = {};

  // Start from the current store minus any existing entries for this extension.
  for (const [panelId, seed] of Object.entries(current)) {
    if (seed.extensionId === owner) continue;
    next[panelId] = seed;
  }

  const seenInExtension = new Set<string>();

  for (const panel of panels ?? []) {
    const panelId = typeof panel?.id === "string" ? panel.id.trim() : "";
    if (!panelId) continue;

    if (seenInExtension.has(panelId)) continue;
    seenInExtension.add(panelId);

    const existing = next[panelId];
    if (existing && existing.extensionId !== owner) {
      const message = `Panel id already contributed by another extension: ${panelId} (existing: ${existing.extensionId}, new: ${owner})`;
      options.onError?.(message);
      return false;
    }

    const titleRaw = typeof panel?.title === "string" ? panel.title.trim() : "";
    const title = titleRaw || panelId;
    const icon = normalizeIcon(panel?.icon);
    const defaultDock = normalizeDockSide(panel?.defaultDock ?? panel?.position);
    next[panelId] = {
      extensionId: owner,
      title,
      ...(icon !== undefined ? { icon } : {}),
      ...(defaultDock ? { defaultDock } : {}),
    };
  }

  writeContributedPanelsSeedStore(storage, next);
  return true;
}

export function removeSeedPanelsForExtension(storage: StorageLike, extensionId: string): void {
  const owner = String(extensionId ?? "").trim();
  if (!owner) return;
  const current = readContributedPanelsSeedStore(storage);
  const next: ContributedPanelsSeedStoreData = {};
  let changed = false;
  for (const [panelId, seed] of Object.entries(current)) {
    if (seed.extensionId === owner) {
      changed = true;
      continue;
    }
    next[panelId] = seed;
  }
  if (!changed) {
    // If the seed store key exists but is already empty (e.g. persisted as `"{}"` by an older
    // client), remove it entirely so uninstall leaves a clean slate.
    if (Object.keys(current).length === 0) {
      safeRemoveItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
    }
    return;
  }
  writeContributedPanelsSeedStore(storage, next);
}

export function clearContributedPanelsSeedStore(storage: StorageLike): void {
  safeRemoveItem(storage, CONTRIBUTED_PANELS_SEED_STORE_KEY);
}

export function getDefaultSeedStoreStorage(): StorageLike | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function seedPanelRegistryFromContributedPanelsSeedStore(
  storage: StorageLike,
  panelRegistry: { registerPanel: (panelId: string, def: any, options?: any) => void },
  options: { onError?: (message: string, err?: unknown) => void } = {},
): void {
  const seeds = readContributedPanelsSeedStore(storage);
  for (const [panelId, seed] of Object.entries(seeds)) {
    try {
      panelRegistry.registerPanel(
        panelId,
        {
          title: seed.title,
          icon: seed.icon ?? null,
          defaultDock: seed.defaultDock ?? "right",
          defaultFloatingRect: { x: 140, y: 140, width: 520, height: 640 },
          source: { kind: "extension", extensionId: seed.extensionId, contributed: true },
        },
        { owner: seed.extensionId },
      );
    } catch (err) {
      options.onError?.(`Failed to seed extension panel: ${panelId}`, err);
    }
  }
}
