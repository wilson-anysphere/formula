import type { CommandRegistry } from "../extensions/commandRegistry.js";

import { COMMAND_RECENTS_STORAGE_KEY, installCommandRecentsTracker, type StorageLike } from "./recents.js";

const DEFAULT_COMMAND_PALETTE_RECENTS_DENYLIST = new Set<string>([
  // The command palette controls its own open behavior; this command is registered as a no-op
  // so other systems (menus, keybindings) can reference it. If we record it, it will always
  // dominate recents.
  "workbench.showCommandPalette",
  // Undo/redo is extremely high frequency during normal editing and not
  // very useful as a "recent command palette command".
  "edit.undo",
  "edit.redo",
  // Extremely frequent commands that should not crowd out more meaningful actions.
  "clipboard.copy",
  "clipboard.cut",
  "clipboard.paste",
  // "Paste Special" variants are also extremely high frequency and typically invoked via menus.
  "clipboard.pasteSpecial",
  "clipboard.pasteSpecial.all",
  "clipboard.pasteSpecial.values",
  "clipboard.pasteSpecial.formulas",
  "clipboard.pasteSpecial.formats",
]);

type TrackerState = {
  refCount: number;
  unsubscribe: () => void;
};

const trackerStateByRegistry = new WeakMap<object, Map<string, TrackerState>>();

/**
 * Keep the Command Palette "RECENT" list in sync with commands executed from anywhere in the app
 * (palette selection, context menus, extensions panel, keybindings, etc).
 */
export function installCommandPaletteRecentsTracking(
  commandRegistry: CommandRegistry,
  storage: StorageLike,
  {
    denylist = DEFAULT_COMMAND_PALETTE_RECENTS_DENYLIST,
    ...options
  }: {
    denylist?: ReadonlySet<string>;
    maxEntries?: number;
    now?: () => number;
    storageKey?: string;
  } = {},
): () => void {
  const storageKey = options.storageKey ?? COMMAND_RECENTS_STORAGE_KEY;

  const map = trackerStateByRegistry.get(commandRegistry) ?? new Map<string, TrackerState>();
  if (!trackerStateByRegistry.has(commandRegistry)) {
    trackerStateByRegistry.set(commandRegistry, map);
  }

  let state = map.get(storageKey);
  if (!state) {
    const unsubscribe = installCommandRecentsTracker(commandRegistry, storage, {
      ...options,
      storageKey,
      ignoreCommandIds: [...denylist],
    });
    state = { refCount: 0, unsubscribe };
    map.set(storageKey, state);
  }

  state.refCount += 1;

  let disposed = false;
  return () => {
    if (disposed) return;
    disposed = true;

    const currentMap = trackerStateByRegistry.get(commandRegistry);
    const current = currentMap?.get(storageKey);
    // The tracker may have been torn down and/or replaced since this disposer was created.
    if (!currentMap || current !== state) return;

    current.refCount -= 1;
    if (current.refCount > 0) return;

    currentMap.delete(storageKey);
    if (currentMap.size === 0) trackerStateByRegistry.delete(commandRegistry);

    try {
      current.unsubscribe();
    } catch {
      // ignore
    }
  };
}
