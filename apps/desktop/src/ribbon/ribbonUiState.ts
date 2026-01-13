export type RibbonUiState = {
  pressedById: Record<string, boolean>;
  labelById: Record<string, string>;
  disabledById: Record<string, boolean>;
  shortcutById: Record<string, string>;
  ariaKeyShortcutsById: Record<string, string>;
};

type Listener = () => void;

let ribbonUiState: RibbonUiState = {
  pressedById: Object.create(null),
  labelById: Object.create(null),
  disabledById: Object.create(null),
  shortcutById: Object.create(null),
  ariaKeyShortcutsById: Object.create(null),
};

const listeners = new Set<Listener>();

export function getRibbonUiStateSnapshot(): RibbonUiState {
  return ribbonUiState;
}

export function subscribeRibbonUiState(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function shallowEqualRecord(a: Record<string, unknown>, b: Record<string, unknown>): boolean {
  if (a === b) return true;
  // Treat prototype changes as meaningful. This allows callers to use prototype chains
  // (e.g. a large baseline object with small per-frame overrides) without losing update
  // notifications when the baseline prototype changes.
  if (Object.getPrototypeOf(a) !== Object.getPrototypeOf(b)) return false;
  const aKeys = Object.keys(a);
  const bKeys = Object.keys(b);
  if (aKeys.length !== bKeys.length) return false;
  for (const key of aKeys) {
    if (a[key] !== b[key]) return false;
  }
  return true;
}

function shallowEqualRibbonUiState(a: RibbonUiState, b: RibbonUiState): boolean {
  return (
    shallowEqualRecord(a.pressedById, b.pressedById) &&
    shallowEqualRecord(a.labelById, b.labelById) &&
    shallowEqualRecord(a.disabledById, b.disabledById) &&
    shallowEqualRecord(a.shortcutById, b.shortcutById) &&
    shallowEqualRecord(a.ariaKeyShortcutsById, b.ariaKeyShortcutsById)
  );
}

export function setRibbonUiState(next: RibbonUiState): void {
  if (shallowEqualRibbonUiState(ribbonUiState, next)) return;
  ribbonUiState = next;
  for (const listener of listeners) listener();
}
