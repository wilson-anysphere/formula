export type RibbonPressedOverrides = Record<string, boolean>;

type Listener = () => void;

let pressedOverrides: RibbonPressedOverrides = Object.create(null);
const listeners = new Set<Listener>();

export function getRibbonPressedOverridesSnapshot(): RibbonPressedOverrides {
  return pressedOverrides;
}

export function subscribeRibbonPressedOverrides(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function shallowEqual(a: RibbonPressedOverrides, b: RibbonPressedOverrides): boolean {
  if (a === b) return true;
  const aKeys = Object.keys(a);
  const bKeys = Object.keys(b);
  if (aKeys.length !== bKeys.length) return false;
  for (const key of aKeys) {
    if (a[key] !== b[key]) return false;
  }
  return true;
}

export function setRibbonPressedOverrides(next: RibbonPressedOverrides): void {
  if (shallowEqual(pressedOverrides, next)) return;
  pressedOverrides = next;
  for (const listener of listeners) listener();
}

