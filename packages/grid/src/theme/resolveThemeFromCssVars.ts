import type { GridTheme } from "./GridTheme";

export const GRID_THEME_CSS_VAR_NAMES: Record<keyof GridTheme, string> = {
  gridBg: "--formula-grid-bg",
  gridLine: "--formula-grid-line",
  headerBg: "--formula-grid-header-bg",
  headerText: "--formula-grid-header-text",
  cellText: "--formula-grid-cell-text",
  errorText: "--formula-grid-error-text",
  selectionFill: "--formula-grid-selection-fill",
  selectionBorder: "--formula-grid-selection-border",
  selectionHandle: "--formula-grid-selection-handle",
  scrollbarTrack: "--formula-grid-scrollbar-track",
  scrollbarThumb: "--formula-grid-scrollbar-thumb",
  freezeLine: "--formula-grid-freeze-line",
  commentIndicator: "--formula-grid-comment-indicator",
  commentIndicatorResolved: "--formula-grid-comment-indicator-resolved",
  remotePresenceDefault: "--formula-grid-remote-presence-default"
};

export function readGridThemeFromCssVars(style: { getPropertyValue: (name: string) => string }): Partial<GridTheme> {
  const partial: Partial<GridTheme> = {};

  for (const [key, cssVar] of Object.entries(GRID_THEME_CSS_VAR_NAMES) as Array<[keyof GridTheme, string]>) {
    const value = style.getPropertyValue(cssVar);
    if (!value) continue;
    const trimmed = value.trim();
    if (!trimmed) continue;
    partial[key] = trimmed;
  }

  return partial;
}

/**
 * Resolve theme tokens from CSS variables on a live DOM element.
 *
 * `getComputedStyle(el).getPropertyValue("--token")` returns the *raw* custom
 * property value. If the token is defined in terms of other CSS variables (e.g.
 * `--formula-grid-bg: var(--app-bg)`), we need an actual CSS property to force
 * variable substitution before passing the value to `canvas.fillStyle`.
 *
 * This helper uses a hidden probe element so the browser resolves nested
 * `var(...)` references for us.
 */
export function resolveGridThemeFromCssVars(element: HTMLElement): Partial<GridTheme> {
  const view = element.ownerDocument?.defaultView;
  if (!view || typeof view.getComputedStyle !== "function") return {};

  const raw = readGridThemeFromCssVars(view.getComputedStyle(element));
  if (Object.keys(raw).length === 0) return raw;

  const probe = element.ownerDocument.createElement("div");
  probe.style.position = "absolute";
  probe.style.width = "0";
  probe.style.height = "0";
  probe.style.overflow = "hidden";
  probe.style.pointerEvents = "none";
  probe.style.visibility = "hidden";
  // Keep the probe from affecting layout/style calculations beyond color parsing.
  probe.style.setProperty("contain", "strict");

  element.appendChild(probe);

  const resolved: Partial<GridTheme> = {};
  try {
    for (const [key, value] of Object.entries(raw) as Array<[keyof GridTheme, string]>) {
      probe.style.backgroundColor = value;
      const computed = view.getComputedStyle(probe).backgroundColor;
      const normalized = computed?.trim();
      resolved[key] = normalized ? normalized : value;
    }
  } finally {
    probe.remove();
  }

  return resolved;
}
