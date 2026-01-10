export const WINDOWING_STATE_VERSION = 1;

function clone(value) {
  return structuredClone(value);
}

function clamp(value, min, max) {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function clampBounds(bounds) {
  return {
    x: typeof bounds.x === "number" && Number.isFinite(bounds.x) ? clamp(bounds.x, -10000, 100000) : null,
    y: typeof bounds.y === "number" && Number.isFinite(bounds.y) ? clamp(bounds.y, -10000, 100000) : null,
    width: clamp(typeof bounds.width === "number" && Number.isFinite(bounds.width) ? bounds.width : 1280, 480, 10000),
    height: clamp(typeof bounds.height === "number" && Number.isFinite(bounds.height) ? bounds.height : 800, 320, 10000),
  };
}

/**
 * @param {{ windows?: Array<{ workbookId: string, workspaceId?: string, bounds?: any, maximized?: boolean }> }} [options]
 */
export function createDefaultWindowingState(options = {}) {
  const windows = Array.isArray(options.windows) ? options.windows : [];

  /** @type {ReturnType<typeof createDefaultWindowingState>} */
  const state = {
    version: WINDOWING_STATE_VERSION,
    nextWindowId: 1,
    windows: [],
    focusedWindowId: null,
  };

  for (const window of windows) {
    if (!window || typeof window !== "object") continue;
    if (typeof window.workbookId !== "string" || window.workbookId.length === 0) continue;
    state.windows.push({
      id: `win${state.nextWindowId++}`,
      workbookId: window.workbookId,
      workspaceId: typeof window.workspaceId === "string" && window.workspaceId.length > 0 ? window.workspaceId : "default",
      bounds: clampBounds(window.bounds ?? {}),
      maximized: Boolean(window.maximized),
    });
  }

  state.focusedWindowId = state.windows[0]?.id ?? null;
  return state;
}

/**
 * @typedef {{ id: string, workbookId: string, workspaceId: string, bounds: { x: number | null, y: number | null, width: number, height: number }, maximized: boolean }} WindowRecord
 */

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} windowId
 */
export function getWindow(state, windowId) {
  return state.windows.find((w) => w.id === windowId) ?? null;
}

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} workbookId
 * @param {{ workspaceId?: string, bounds?: any, maximized?: boolean, focus?: boolean }} [options]
 */
export function openWorkbookWindow(state, workbookId, options = {}) {
  if (typeof workbookId !== "string" || workbookId.length === 0) {
    throw new Error("workbookId must be a non-empty string");
  }

  const next = clone(state);
  const id = `win${next.nextWindowId++}`;
  next.windows.push({
    id,
    workbookId,
    workspaceId: typeof options.workspaceId === "string" && options.workspaceId.length > 0 ? options.workspaceId : "default",
    bounds: clampBounds(options.bounds ?? {}),
    maximized: Boolean(options.maximized),
  });

  if (options.focus ?? true) next.focusedWindowId = id;
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} windowId
 */
export function closeWindow(state, windowId) {
  const existing = getWindow(state, windowId);
  if (!existing) return state;

  const next = clone(state);
  next.windows = next.windows.filter((w) => w.id !== windowId);

  if (next.focusedWindowId === windowId) {
    next.focusedWindowId = next.windows.at(-1)?.id ?? next.windows[0]?.id ?? null;
  }

  return next;
}

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} windowId
 */
export function focusWindow(state, windowId) {
  if (!getWindow(state, windowId)) return state;
  if (state.focusedWindowId === windowId) return state;

  const next = clone(state);
  next.focusedWindowId = windowId;
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} windowId
 * @param {{ x?: number | null, y?: number | null, width?: number, height?: number }} bounds
 */
export function setWindowBounds(state, windowId, bounds) {
  const existing = getWindow(state, windowId);
  if (!existing) return state;

  const next = clone(state);
  const idx = next.windows.findIndex((w) => w.id === windowId);
  next.windows[idx] = { ...next.windows[idx], bounds: clampBounds({ ...existing.bounds, ...bounds }) };
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultWindowingState>} state
 * @param {string} windowId
 * @param {boolean} maximized
 */
export function setWindowMaximized(state, windowId, maximized) {
  const existing = getWindow(state, windowId);
  if (!existing) return state;

  const next = clone(state);
  const idx = next.windows.findIndex((w) => w.id === windowId);
  next.windows[idx] = { ...next.windows[idx], maximized: Boolean(maximized) };
  return next;
}

