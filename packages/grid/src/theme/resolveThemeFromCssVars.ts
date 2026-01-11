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

function splitCssVarArguments(inner: string): { name: string; fallback: string | null } | null {
  const trimmed = inner.trim();
  if (!trimmed.startsWith("--")) return null;

  let depth = 0;
  let commaIndex = -1;
  for (let i = 0; i < trimmed.length; i++) {
    const ch = trimmed[i];
    if (ch === "(") depth += 1;
    else if (ch === ")") depth = Math.max(0, depth - 1);
    else if (ch === "," && depth === 0) {
      commaIndex = i;
      break;
    }
  }

  if (commaIndex === -1) return { name: trimmed, fallback: null };

  return {
    name: trimmed.slice(0, commaIndex).trim(),
    fallback: trimmed.slice(commaIndex + 1).trim() || null
  };
}

function parseCssVarFunction(value: string): { name: string; fallback: string | null } | null {
  const trimmed = value.trim();
  if (!trimmed.startsWith("var(") || !trimmed.endsWith(")")) return null;
  const inner = trimmed.slice(4, -1);
  return splitCssVarArguments(inner);
}

/**
 * Best-effort resolver for `var(--token)` values from computed custom properties.
 *
 * This intentionally handles the common case where a theme token references a
 * single other custom property (optionally with a fallback), e.g.
 * `--formula-grid-bg: var(--app-bg, #fff)`.
 *
 * It does *not* attempt to fully evaluate arbitrary CSS expressions.
 */
export function resolveCssVarValue(
  value: string,
  style: { getPropertyValue: (name: string) => string },
  options?: { maxDepth?: number }
): string {
  const maxDepth = options?.maxDepth ?? 10;
  let current = value.trim();
  const seen = new Set<string>();

  for (let depth = 0; depth < maxDepth; depth++) {
    const parsed = parseCssVarFunction(current);
    if (!parsed) break;

    const { name, fallback } = parsed;
    if (seen.has(name)) {
      current = fallback ?? "";
      continue;
    }
    seen.add(name);

    const next = style.getPropertyValue(name).trim();
    if (next) {
      current = next;
      continue;
    }

    current = fallback ?? "";
  }

  return current;
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

  const computedStyle = view.getComputedStyle(element);
  const raw = readGridThemeFromCssVars(computedStyle);
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
      const resolvedValue = value.includes("var(") ? resolveCssVarValue(value, computedStyle) : value;
      probe.style.backgroundColor = resolvedValue;
      const computed = view.getComputedStyle(probe).backgroundColor;
      const normalized = computed?.trim();
      resolved[key] = normalized ? normalized : resolvedValue;
    }
  } finally {
    probe.remove();
  }

  return resolved;
}
