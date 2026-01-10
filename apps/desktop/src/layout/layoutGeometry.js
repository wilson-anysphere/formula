import { SPLIT_DIRECTIONS } from "./constants.js";

function clamp(value, min, max) {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function ensureFiniteNumber(value, fallback = 0) {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function rect(x, y, width, height) {
  return { x, y, width, height };
}

function zoneVisible(zone) {
  return Boolean(zone && !zone.collapsed && Array.isArray(zone.panels) && zone.panels.length > 0);
}

/**
 * Computes the rectangles for dock zones and the remaining content area.
 *
 * This is purely geometric and UI-agnostic: callers can use it to position a grid view,
 * dock hosts, and split panes.
 *
 * @param {{ docks: Record<"left" | "right" | "bottom", { size: number, collapsed: boolean, panels: string[] }> }} layout
 * @param {{ width: number, height: number }} viewport
 */
export function computeDockLayoutRects(layout, viewport) {
  const width = Math.max(0, ensureFiniteNumber(viewport.width));
  const height = Math.max(0, ensureFiniteNumber(viewport.height));

  const leftZone = layout.docks.left;
  const rightZone = layout.docks.right;
  const bottomZone = layout.docks.bottom;

  const leftSize = zoneVisible(leftZone) ? clamp(ensureFiniteNumber(leftZone.size), 0, width) : 0;
  const rightSize = zoneVisible(rightZone) ? clamp(ensureFiniteNumber(rightZone.size), 0, width - leftSize) : 0;
  const bottomSize = zoneVisible(bottomZone)
    ? clamp(ensureFiniteNumber(bottomZone.size), 0, height)
    : 0;

  const contentWidth = Math.max(0, width - leftSize - rightSize);
  const contentHeight = Math.max(0, height - bottomSize);

  const docks = {
    left: leftSize > 0 ? rect(0, 0, leftSize, contentHeight) : null,
    right: rightSize > 0 ? rect(width - rightSize, 0, rightSize, contentHeight) : null,
    bottom: bottomSize > 0 ? rect(0, height - bottomSize, width, bottomSize) : null,
  };

  return {
    viewport: rect(0, 0, width, height),
    docks,
    content: rect(leftSize, 0, contentWidth, contentHeight),
  };
}

/**
 * Computes split-pane rectangles within a content rect.
 *
 * @param {{ direction: "none" | "vertical" | "horizontal", ratio: number }} splitView
 * @param {{ x: number, y: number, width: number, height: number }} contentRect
 * @param {{ gutter?: number }} [options]
 */
export function computeSplitViewRects(splitView, contentRect, options = {}) {
  const direction = SPLIT_DIRECTIONS.includes(splitView.direction) ? splitView.direction : "none";
  const gutter = Math.max(0, ensureFiniteNumber(options.gutter, 4));
  const ratio = clamp(ensureFiniteNumber(splitView.ratio, 0.5), 0.1, 0.9);

  if (direction === "none") {
    return {
      direction,
      primary: contentRect,
      secondary: null,
      gutter: null,
    };
  }

  if (direction === "vertical") {
    const available = Math.max(0, contentRect.width - gutter);
    const primaryWidth = Math.round(available * ratio);
    const secondaryWidth = Math.max(0, available - primaryWidth);

    const primary = rect(contentRect.x, contentRect.y, primaryWidth, contentRect.height);
    const gutterRect = rect(contentRect.x + primaryWidth, contentRect.y, gutter, contentRect.height);
    const secondary = rect(
      contentRect.x + primaryWidth + gutter,
      contentRect.y,
      secondaryWidth,
      contentRect.height,
    );

    return { direction, primary, secondary, gutter: gutterRect };
  }

  // horizontal
  const available = Math.max(0, contentRect.height - gutter);
  const primaryHeight = Math.round(available * ratio);
  const secondaryHeight = Math.max(0, available - primaryHeight);

  const primary = rect(contentRect.x, contentRect.y, contentRect.width, primaryHeight);
  const gutterRect = rect(contentRect.x, contentRect.y + primaryHeight, contentRect.width, gutter);
  const secondary = rect(
    contentRect.x,
    contentRect.y + primaryHeight + gutter,
    contentRect.width,
    secondaryHeight,
  );

  return { direction, primary, secondary, gutter: gutterRect };
}

/**
 * Convenience helper that computes both dock + split rects.
 *
 * @param {{ docks: Record<"left" | "right" | "bottom", { size: number, collapsed: boolean, panels: string[] }>, splitView: { direction: "none" | "vertical" | "horizontal", ratio: number } }} layout
 * @param {{ width: number, height: number }} viewport
 * @param {{ gutter?: number }} [options]
 */
export function computeWorkspaceRects(layout, viewport, options = {}) {
  const dockRects = computeDockLayoutRects(layout, viewport);
  const splitRects = computeSplitViewRects(layout.splitView, dockRects.content, options);

  return {
    ...dockRects,
    split: splitRects,
  };
}
