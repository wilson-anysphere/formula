/**
 * Minimal runtime helpers for safely accessing Tauri's `core.invoke` without a hard
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
 * Cache of wrappers that preserve invoke method binding + argument count without mutating
 * the injected `__TAURI__` object.
 *
 * @type {WeakMap<object, { invoke: Function, bound: TauriInvoke }>}
 */
const boundInvokeCache = new WeakMap();

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
  if (typeof invoke !== "function") return null;
  if (core && (typeof core === "object" || typeof core === "function")) {
    const cached = boundInvokeCache.get(core);
    if (cached?.invoke === invoke) return cached.bound;
    const bound = /** @type {TauriInvoke} */ (function (cmd, args) {
      if (arguments.length < 2) {
        return invoke.call(core, cmd);
      }
      return invoke.call(core, cmd, args);
    });
    boundInvokeCache.set(core, { invoke, bound });
    return bound;
  }

  // Extremely defensive fallback: bind without caching when `core` is not a WeakMap key.
  return /** @type {TauriInvoke} */ (function (cmd, args) {
    if (arguments.length < 2) {
      return invoke.call(core, cmd);
    }
    return invoke.call(core, cmd, args);
  });
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
