export const KEYBINDING_BARRIER_ATTRIBUTE = "data-keybinding-barrier";

/**
 * Mark an element as a keybinding barrier.
 *
 * Global keyboard shortcut handlers (KeybindingService, spreadsheet-global listeners, etc)
 * should ignore key events whose composed path includes an element with this attribute.
 *
 * @template {Element} T
 * @param {T} el
 * @returns {T}
 */
export function markKeybindingBarrier(el) {
  el.setAttribute(KEYBINDING_BARRIER_ATTRIBUTE, "true");
  return el;
}

/**
 * Returns true if the event originated from within a keybinding barrier.
 *
 * @param {Event | null | undefined} event
 * @returns {boolean}
 */
export function isEventWithinKeybindingBarrier(event) {
  if (!event) return false;

  const isElement = (node) => {
    if (!node || (typeof node !== "object" && typeof node !== "function")) return false;
    // Guard against Node environments without a DOM global.
    if (typeof Element !== "undefined") return node instanceof Element;
    // Fallback for jsdom-like globals where `Element` isn't hoisted onto `globalThis`.
    return (
      // nodeType 1 === ELEMENT_NODE
      node.nodeType === 1 &&
      typeof node.hasAttribute === "function" &&
      typeof node.getAttribute === "function"
    );
  };

  // `composedPath()` is the most reliable way to detect barriers across shadow DOM boundaries.
  if (typeof event.composedPath === "function") {
    try {
      const path = event.composedPath();
      for (const node of path) {
        if (!isElement(node)) continue;
        if (node.hasAttribute(KEYBINDING_BARRIER_ATTRIBUTE)) return true;
      }
      return false;
    } catch {
      // Fall back to DOM walking below.
    }
  }

  const target = event.target;
  if (!isElement(target)) return false;
  try {
    return Boolean(target.closest?.(`[${KEYBINDING_BARRIER_ATTRIBUTE}]`));
  } catch {
    return false;
  }
}
