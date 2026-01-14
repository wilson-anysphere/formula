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
import { createDefaultLayout } from "./layoutState.js";
import { MAX_GRID_ZOOM, MIN_GRID_ZOOM } from "@formula/grid/node";

const hasOwn = (obj, key) => Object.prototype.hasOwnProperty.call(obj, key);

function panelRegistryHas(panelRegistry, panelId) {
  if (!panelRegistry) return true;
  const id = String(panelId);
  if (typeof panelRegistry.has === "function") return panelRegistry.has(id);
  if (typeof panelRegistry.hasPanel === "function") return panelRegistry.hasPanel(id);
  return hasOwn(panelRegistry, id);
}

function clampNumber(value, { min, max, fallback }) {
  if (typeof value !== "number" || Number.isNaN(value)) return fallback;
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function normalizeDockZone(rawZone, side, { panelRegistry }) {
  const defaultSize = DEFAULT_DOCK_SIZES[side];

  /** @type {{ size: number, collapsed: boolean, panels: string[], active: string | null }} */
  const zone = {
    size: defaultSize,
    collapsed: false,
    panels: [],
    active: null,
  };

  if (!rawZone || typeof rawZone !== "object") return zone;

  zone.size = clampNumber(rawZone.size, { min: 120, max: 1200, fallback: defaultSize });
  zone.collapsed = Boolean(rawZone.collapsed);

  if (Array.isArray(rawZone.panels)) {
    const panels = rawZone.panels.filter((id) => typeof id === "string" && id.length > 0);
    zone.panels = panelRegistry ? panels.filter((id) => panelRegistryHas(panelRegistry, id)) : panels;
  }

  zone.active = typeof rawZone.active === "string" ? rawZone.active : null;
  if (zone.active && !zone.panels.includes(zone.active)) zone.active = null;
  if (!zone.active && zone.panels.length > 0) zone.active = zone.panels[0];

  return zone;
}

function normalizeFloating(rawFloating, { panelRegistry }) {
  /** @type {Record<string, { x: number, y: number, width: number, height: number, minimized: boolean }>} */
  const floating = {};

  if (!rawFloating || typeof rawFloating !== "object") return floating;

  for (const [panelId, rawPanel] of Object.entries(rawFloating)) {
    if (typeof panelId !== "string" || panelId.length === 0) continue;
    if (panelRegistry && !panelRegistryHas(panelRegistry, panelId)) continue;
    if (!rawPanel || typeof rawPanel !== "object") continue;

    const x = clampNumber(rawPanel.x, { min: -10000, max: 100000, fallback: DEFAULT_FLOATING_RECT.x });
    const y = clampNumber(rawPanel.y, { min: -10000, max: 100000, fallback: DEFAULT_FLOATING_RECT.y });
    const width = clampNumber(rawPanel.width, { min: 160, max: 2000, fallback: DEFAULT_FLOATING_RECT.width });
    const height = clampNumber(rawPanel.height, { min: 120, max: 2000, fallback: DEFAULT_FLOATING_RECT.height });

    floating[panelId] = {
      x,
      y,
      width,
      height,
      minimized: Boolean(rawPanel.minimized),
    };
  }

  return floating;
}

function normalizeSplitView(rawSplitView, { primarySheetId }) {
  /** @type {{ direction: "none" | "vertical" | "horizontal", ratio: number, activePane: "primary" | "secondary", panes: { primary: { sheetId: string | null, scrollX: number, scrollY: number, zoom: number }, secondary: { sheetId: string | null, scrollX: number, scrollY: number, zoom: number } } }} */
  const splitView = {
    direction: "none",
    ratio: DEFAULT_SPLIT_RATIO,
    activePane: DEFAULT_ACTIVE_SPLIT_PANE,
    panes: {
      primary: {
        sheetId: primarySheetId ?? null,
        scrollX: 0,
        scrollY: 0,
        zoom: DEFAULT_PANE_ZOOM,
      },
      secondary: {
        sheetId: primarySheetId ?? null,
        scrollX: 0,
        scrollY: 0,
        zoom: DEFAULT_PANE_ZOOM,
      },
    },
  };

  if (!rawSplitView || typeof rawSplitView !== "object") return splitView;

  splitView.direction = SPLIT_DIRECTIONS.includes(rawSplitView.direction) ? rawSplitView.direction : "none";
  splitView.ratio = clampNumber(rawSplitView.ratio, { min: 0.1, max: 0.9, fallback: DEFAULT_SPLIT_RATIO });
  splitView.activePane = SPLIT_PANES.includes(rawSplitView.activePane) ? rawSplitView.activePane : DEFAULT_ACTIVE_SPLIT_PANE;

  const panes = rawSplitView.panes;
  if (panes && typeof panes === "object") {
    for (const paneId of SPLIT_PANES) {
      const rawPane = panes[paneId];
      if (!rawPane || typeof rawPane !== "object") continue;
      splitView.panes[paneId] = {
        sheetId: typeof rawPane.sheetId === "string" ? rawPane.sheetId : splitView.panes[paneId].sheetId,
        scrollX: clampNumber(rawPane.scrollX, { min: -1e12, max: 1e12, fallback: splitView.panes[paneId].scrollX }),
        scrollY: clampNumber(rawPane.scrollY, { min: -1e12, max: 1e12, fallback: splitView.panes[paneId].scrollY }),
        zoom: clampNumber(rawPane.zoom, { min: MIN_GRID_ZOOM, max: MAX_GRID_ZOOM, fallback: splitView.panes[paneId].zoom }),
      };
    }
  }

  return splitView;
}

function dedupePanelLocations(layout) {
  /** @type {Set<string>} */
  const seen = new Set();

  for (const side of DOCK_SIDES) {
    const zone = layout.docks[side];
    zone.panels = zone.panels.filter((id) => {
      if (seen.has(id)) return false;
      seen.add(id);
      return true;
    });
    if (zone.active && !zone.panels.includes(zone.active)) zone.active = zone.panels[0] ?? null;
  }

  for (const panelId of Object.keys(layout.floating)) {
    if (seen.has(panelId)) {
      delete layout.floating[panelId];
      continue;
    }
    seen.add(panelId);
  }
}

/**
 * Normalizes a potentially-partial/unknown layout blob into a fully-formed state
 * that is safe to render and safe to persist.
 *
 * @param {unknown} raw
 * @param {{ panelRegistry?: Record<string, unknown>, primarySheetId?: string | null } } [options]
 */
export function normalizeLayout(raw, options = {}) {
  const { panelRegistry, primarySheetId } = options;

  if (!raw || typeof raw !== "object") {
    return createDefaultLayout({ primarySheetId });
  }

  const version = raw.version === LAYOUT_STATE_VERSION ? raw.version : LAYOUT_STATE_VERSION;

  /** @type {ReturnType<typeof createDefaultLayout>} */
  const layout = {
    version,
    docks: {
      left: normalizeDockZone(raw.docks?.left, "left", { panelRegistry }),
      right: normalizeDockZone(raw.docks?.right, "right", { panelRegistry }),
      bottom: normalizeDockZone(raw.docks?.bottom, "bottom", { panelRegistry }),
    },
    floating: normalizeFloating(raw.floating, { panelRegistry }),
    splitView: normalizeSplitView(raw.splitView, { primarySheetId }),
  };

  dedupePanelLocations(layout);

  return layout;
}
