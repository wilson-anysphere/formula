/**
 * Minimal runtime helpers for safely accessing `__TAURI__.core.invoke` without a hard
 * dependency on `@tauri-apps/api`.
 *
 * This module is intentionally JavaScript so it can be consumed by other `.js` sources
 * (and node:test suites) without requiring TypeScript execution support.
 *
 * @typedef {(cmd: string, args?: any) => Promise<any>} TauriInvoke
 */

/**
 * @param {any} obj
 * @param {string} prop
 * @returns {any | undefined}
 */
function safeGetProp(obj, prop) {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

/**
 * @returns {any | null}
 */
function getTauriGlobalOrNull() {
  try {
    return globalThis.__TAURI__ ?? null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so best-effort callsites can fall back cleanly.
    return null;
  }
}

/**
 * @returns {TauriInvoke | null}
 */
export function getTauriInvokeOrNull() {
  const tauri = getTauriGlobalOrNull();
  const core = safeGetProp(tauri, "core");
  const invoke = /** @type {unknown} */ (safeGetProp(core, "invoke"));
  return typeof invoke === "function" ? /** @type {TauriInvoke} */ (invoke) : null;
}

/**
 * @returns {TauriInvoke}
 */
export function getTauriInvokeOrThrow() {
  const invoke = getTauriInvokeOrNull();
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export function hasTauriInvoke() {
  return getTauriInvokeOrNull() != null;
}

