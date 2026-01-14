import {
  DEFAULT_DOCK_SIZES,
  DEFAULT_FLOATING_RECT,
  DEFAULT_ACTIVE_SPLIT_PANE,
  DEFAULT_PANE_ZOOM,
  DEFAULT_SPLIT_RATIO,
  DOCK_SIDES,
  LAYOUT_STATE_VERSION,
  SPLIT_DIRECTIONS,
  SPLIT_PANES,
} from "./constants.js";
import { MAX_GRID_ZOOM, MIN_GRID_ZOOM } from "@formula/grid/node";

function clone(value) {
  return structuredClone(value);
}

function clamp(value, min, max) {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function clampRect(rect) {
  return {
    x: clamp(rect.x, -10000, 100000),
    y: clamp(rect.y, -10000, 100000),
    width: clamp(rect.width, 160, 2000),
    height: clamp(rect.height, 120, 2000),
  };
}

function removeFromDockZone(zone, panelId) {
  const next = clone(zone);
  next.panels = next.panels.filter((id) => id !== panelId);
  if (next.active === panelId) next.active = next.panels[0] ?? null;
  return next;
}

function removePanelFromAllDocks(layout, panelId) {
  for (const side of DOCK_SIDES) {
    layout.docks[side] = removeFromDockZone(layout.docks[side], panelId);
  }
}

function ensureDockSide(side) {
  if (!DOCK_SIDES.includes(side)) {
    throw new Error(`Invalid dock side: ${side}`);
  }
}

function ensureSplitDirection(direction) {
  if (!SPLIT_DIRECTIONS.includes(direction)) {
    throw new Error(`Invalid split direction: ${direction}`);
  }
}

function ensureSplitPane(pane) {
  if (!SPLIT_PANES.includes(pane)) {
    throw new Error(`Invalid split pane: ${pane}`);
  }
}

function panelRegistryGet(panelRegistry, panelId) {
  if (!panelRegistry) return undefined;
  const id = String(panelId);
  if (typeof panelRegistry.get === "function") return panelRegistry.get(id);
  if (typeof panelRegistry.getPanel === "function") return panelRegistry.getPanel(id);
  return panelRegistry[id];
}

/**
 * @param {{ primarySheetId?: string | null }} [options]
 */
export function createDefaultLayout(options = {}) {
  const primarySheetId = options.primarySheetId ?? null;

  return {
    version: LAYOUT_STATE_VERSION,
    docks: {
      left: { size: DEFAULT_DOCK_SIZES.left, collapsed: false, panels: [], active: null },
      right: { size: DEFAULT_DOCK_SIZES.right, collapsed: false, panels: [], active: null },
      bottom: { size: DEFAULT_DOCK_SIZES.bottom, collapsed: false, panels: [], active: null },
    },
    floating: {},
    splitView: {
      direction: "none",
      ratio: DEFAULT_SPLIT_RATIO,
      activePane: DEFAULT_ACTIVE_SPLIT_PANE,
      panes: {
        primary: { sheetId: primarySheetId, scrollX: 0, scrollY: 0, zoom: DEFAULT_PANE_ZOOM },
        secondary: { sheetId: primarySheetId, scrollX: 0, scrollY: 0, zoom: DEFAULT_PANE_ZOOM },
      },
    },
  };
}

/**
 * Returns the current placement of a panel.
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 */
export function getPanelPlacement(layout, panelId) {
  for (const side of DOCK_SIDES) {
    if (layout.docks[side].panels.includes(panelId)) {
      return { kind: "docked", side };
    }
  }

  if (layout.floating[panelId]) return { kind: "floating" };

  return { kind: "closed" };
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {{ panelRegistry?: any }} [options]
 */
export function openPanel(layout, panelId, options = {}) {
  const placement = getPanelPlacement(layout, panelId);
  if (placement.kind === "docked") {
    return activateDockedPanel(layout, panelId, placement.side);
  }
  if (placement.kind === "floating") return layout;

  const defaults = panelRegistryGet(options.panelRegistry, panelId);
  if (defaults?.defaultDock) {
    return dockPanel(layout, panelId, defaults.defaultDock, { activate: true });
  }

  return floatPanel(layout, panelId, defaults?.defaultFloatingRect ?? DEFAULT_FLOATING_RECT, { activate: true });
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 */
export function closePanel(layout, panelId) {
  const next = clone(layout);
  removePanelFromAllDocks(next, panelId);
  delete next.floating[panelId];
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {"left" | "right" | "bottom"} side
 * @param {{ activate?: boolean, insert?: "start" | "end" }} [options]
 */
export function dockPanel(layout, panelId, side, options = {}) {
  ensureDockSide(side);
  const next = clone(layout);

  removePanelFromAllDocks(next, panelId);
  delete next.floating[panelId];

  const zone = clone(next.docks[side]);
  zone.panels = zone.panels.filter((id) => id !== panelId);
  if (options.insert === "start") zone.panels.unshift(panelId);
  else zone.panels.push(panelId);
  if (options.activate ?? true) zone.active = panelId;
  zone.collapsed = false;

  next.docks[side] = zone;
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {{ x: number, y: number, width: number, height: number }} rect
 * @param {{ activate?: boolean }} [options]
 */
export function floatPanel(layout, panelId, rect, options = {}) {
  const next = clone(layout);

  removePanelFromAllDocks(next, panelId);

  const nextRect = clampRect(rect);

  next.floating[panelId] = {
    x: nextRect.x,
    y: nextRect.y,
    width: nextRect.width,
    height: nextRect.height,
    minimized: false,
  };

  if (options.activate ?? true) {
    // no-op: floating panels are always "active" independently
  }

  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {boolean} minimized
 */
export function setFloatingPanelMinimized(layout, panelId, minimized) {
  const existing = layout.floating[panelId];
  if (!existing) return layout;

  const next = clone(layout);
  next.floating[panelId] = { ...existing, minimized: Boolean(minimized) };
  return next;
}

/**
 * Update the position/size of a floating panel (e.g. drag/resize).
 *
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {{ x: number, y: number, width: number, height: number }} rect
 */
export function setFloatingPanelRect(layout, panelId, rect) {
  const existing = layout.floating[panelId];
  if (!existing) return layout;

  const next = clone(layout);
  const nextRect = clampRect(rect);
  next.floating[panelId] = { ...existing, ...nextRect };
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"left" | "right" | "bottom"} side
 * @param {boolean} collapsed
 */
export function setDockCollapsed(layout, side, collapsed) {
  ensureDockSide(side);
  const next = clone(layout);
  next.docks[side] = { ...next.docks[side], collapsed: Boolean(collapsed) };
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"left" | "right" | "bottom"} side
 * @param {number} sizePx
 */
export function setDockSize(layout, side, sizePx) {
  ensureDockSide(side);
  const next = clone(layout);
  const size = clamp(sizePx, 120, 1200);
  next.docks[side] = { ...next.docks[side], size };
  return next;
}

/**
 * Snaps a floating panel into a dock zone when it is close enough to an edge.
 * This is UI-agnostic: the caller provides the viewport and threshold.
 *
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {{ width: number, height: number }} viewport
 * @param {{ threshold?: number }} [options]
 */
export function snapFloatingPanel(layout, panelId, viewport, options = {}) {
  const floating = layout.floating[panelId];
  if (!floating) return layout;

  const threshold = Math.max(0, options.threshold ?? 24);
  const viewportWidth = typeof viewport.width === "number" ? viewport.width : NaN;
  const viewportHeight = typeof viewport.height === "number" ? viewport.height : NaN;
  if (!Number.isFinite(viewportWidth) || !Number.isFinite(viewportHeight)) return layout;

  const distances = /** @type {Array<{ side: "left" | "right" | "bottom", distance: number }>} */ ([
    { side: "left", distance: floating.x },
    { side: "right", distance: viewportWidth - (floating.x + floating.width) },
    { side: "bottom", distance: viewportHeight - (floating.y + floating.height) },
  ]);

  const candidate = distances
    .map((d) => ({ ...d, abs: Math.abs(d.distance) }))
    .filter((d) => Number.isFinite(d.abs) && d.abs <= threshold)
    .sort((a, b) => a.abs - b.abs)[0];

  if (!candidate) return layout;

  return dockPanel(layout, panelId, candidate.side, { activate: true });
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {string} panelId
 * @param {"left" | "right" | "bottom"} side
 */
export function activateDockedPanel(layout, panelId, side) {
  ensureDockSide(side);
  if (!layout.docks[side].panels.includes(panelId)) return layout;

  const next = clone(layout);
  next.docks[side] = { ...next.docks[side], active: panelId, collapsed: false };
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"none" | "vertical" | "horizontal"} direction
 * @param {number} [ratio]
 */
export function setSplitDirection(layout, direction, ratio) {
  ensureSplitDirection(direction);

  const next = clone(layout);
  next.splitView.direction = direction;
  if (typeof ratio === "number") {
    next.splitView.ratio = clamp(ratio, 0.1, 0.9);
  }

  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {number} ratio
 */
export function setSplitRatio(layout, ratio) {
  // Split ratio is updated at high frequency (splitter drag). Avoid `structuredClone(layout)`
  // here: LayoutController normalizes layouts on every commit anyway, and the split ratio update
  // only needs a shallow copy with an updated `splitView` object.
  const clamped = clamp(ratio, 0.1, 0.9);
  if (layout?.splitView?.ratio === clamped) return layout;
  return { ...layout, splitView: { ...layout.splitView, ratio: clamped } };
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"primary" | "secondary"} pane
 */
export function setActiveSplitPane(layout, pane) {
  ensureSplitPane(pane);
  if (layout.splitView.activePane === pane) return layout;

  const next = clone(layout);
  next.splitView.activePane = pane;
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"primary" | "secondary"} pane
 * @param {string} sheetId
 */
export function setSplitPaneSheet(layout, pane, sheetId) {
  ensureSplitPane(pane);

  const next = clone(layout);
  next.splitView.panes[pane] = { ...next.splitView.panes[pane], sheetId };
  return next;
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"primary" | "secondary"} pane
 * @param {{ scrollX: number, scrollY: number }} scroll
 */
export function setSplitPaneScroll(layout, pane, scroll) {
  ensureSplitPane(pane);
  const existing = layout.splitView.panes[pane];
  const scrollX = scroll.scrollX;
  const scrollY = scroll.scrollY;
  if (existing.scrollX === scrollX && existing.scrollY === scrollY) return layout;
  return {
    ...layout,
    splitView: {
      ...layout.splitView,
      panes: {
        ...layout.splitView.panes,
        [pane]: { ...existing, scrollX, scrollY },
      },
    },
  };
}

/**
 * @param {ReturnType<typeof createDefaultLayout>} layout
 * @param {"primary" | "secondary"} pane
 * @param {number} zoom
 */
export function setSplitPaneZoom(layout, pane, zoom) {
  ensureSplitPane(pane);
  const existing = layout.splitView.panes[pane];
  const value = typeof zoom === "number" && Number.isFinite(zoom) ? zoom : DEFAULT_PANE_ZOOM;
  const clamped = clamp(value, MIN_GRID_ZOOM, MAX_GRID_ZOOM);
  if (existing.zoom === clamped) return layout;
  return {
    ...layout,
    splitView: {
      ...layout.splitView,
      panes: {
        ...layout.splitView.panes,
        [pane]: { ...existing, zoom: clamped },
      },
    },
  };
}
