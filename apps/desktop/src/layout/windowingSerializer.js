import { createDefaultWindowingState, WINDOWING_STATE_VERSION } from "./windowingState.js";

function clamp(value, min, max) {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function normalizeWindowRecord(raw, idFallback) {
  const workbookId = typeof raw?.workbookId === "string" ? raw.workbookId : "";
  if (!workbookId) return null;

  const id = typeof raw?.id === "string" && raw.id.length > 0 ? raw.id : idFallback;
  const workspaceId = typeof raw?.workspaceId === "string" && raw.workspaceId.length > 0 ? raw.workspaceId : "default";
  const bounds = raw?.bounds && typeof raw.bounds === "object" ? raw.bounds : {};

  return {
    id,
    workbookId,
    workspaceId,
    bounds: {
      x: typeof bounds.x === "number" && Number.isFinite(bounds.x) ? clamp(bounds.x, -10000, 100000) : null,
      y: typeof bounds.y === "number" && Number.isFinite(bounds.y) ? clamp(bounds.y, -10000, 100000) : null,
      width: clamp(typeof bounds.width === "number" && Number.isFinite(bounds.width) ? bounds.width : 1280, 480, 10000),
      height: clamp(typeof bounds.height === "number" && Number.isFinite(bounds.height) ? bounds.height : 800, 320, 10000),
    },
    maximized: Boolean(raw?.maximized),
  };
}

function normalizeWindowingState(raw) {
  if (!raw || typeof raw !== "object") return createDefaultWindowingState();

  const windowsRaw = Array.isArray(raw.windows) ? raw.windows : [];
  const windows = [];
  const seen = new Set();

  const rawNextWindowId =
    typeof raw.nextWindowId === "number" && Number.isFinite(raw.nextWindowId) ? Math.floor(raw.nextWindowId) : 1;

  let idCounter = 1;
  let maxNumericId = 0;

  function bumpMaxNumeric(id) {
    const match = /^win(\d+)$/.exec(id);
    if (!match) return;
    const value = Number(match[1]);
    if (Number.isInteger(value) && value > maxNumericId) maxNumericId = value;
  }

  for (const entry of windowsRaw) {
    const normalized = normalizeWindowRecord(entry, "");
    if (!normalized) continue;

    let id = normalized.id;
    if (!id || seen.has(id)) {
      while (seen.has(`win${idCounter}`)) idCounter += 1;
      id = `win${idCounter++}`;
    }

    normalized.id = id;

    if (seen.has(id)) continue;
    seen.add(id);
    bumpMaxNumeric(id);
    windows.push(normalized);
  }

  // Ensure future generated ids don't collide with existing `winN` ids.
  const nextWindowId = Math.max(1, rawNextWindowId, maxNumericId + 1, idCounter);

  const focusedWindowId =
    typeof raw.focusedWindowId === "string" && windows.some((w) => w.id === raw.focusedWindowId)
      ? raw.focusedWindowId
      : windows[0]?.id ?? null;

  return {
    version: WINDOWING_STATE_VERSION,
    nextWindowId,
    windows,
    focusedWindowId,
  };
}

/**
 * @param {unknown} state
 */
export function serializeWindowingState(state) {
  return JSON.stringify(normalizeWindowingState(state));
}

/**
 * @param {string} json
 */
export function deserializeWindowingState(json) {
  try {
    return normalizeWindowingState(JSON.parse(json));
  } catch {
    return createDefaultWindowingState();
  }
}
