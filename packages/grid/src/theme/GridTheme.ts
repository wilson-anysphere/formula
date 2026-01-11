export interface GridTheme {
  /** Background color for the entire grid surface. */
  gridBg: string;
  /** Grid line (cell border) color. */
  gridLine: string;
  /** Background color for the header row/col (when enabled via frozen rows/cols). */
  headerBg: string;
  /** Text color for header cells. */
  headerText: string;
  /** Default text color for regular cells (unless a cell style overrides it). */
  cellText: string;
  /** Text color used for error strings (values starting with `#`, unless overridden). */
  errorText: string;
  /** Fill color for selection rectangles. */
  selectionFill: string;
  /** Stroke color for selection rectangles. */
  selectionBorder: string;
  /** Fill color for the selection handle (bottom-right square). */
  selectionHandle: string;
  /** Scrollbar track background. */
  scrollbarTrack: string;
  /** Scrollbar thumb background. */
  scrollbarThumb: string;
  /** Freeze pane divider line color. */
  freezeLine: string;
  /** Comment indicator (unresolved) triangle fill. */
  commentIndicator: string;
  /** Comment indicator (resolved) triangle fill. */
  commentIndicatorResolved: string;
  /** Fallback color for remote presence overlays when no per-presence color is provided. */
  remotePresenceDefault: string;
}

export const DEFAULT_GRID_THEME: GridTheme = {
  gridBg: "#ffffff",
  gridLine: "#e6e6e6",
  headerBg: "#f8fafc",
  headerText: "#0f172a",
  cellText: "#111111",
  errorText: "#cc0000",
  selectionFill: "rgba(14, 101, 235, 0.12)",
  selectionBorder: "#0e65eb",
  selectionHandle: "#0e65eb",
  scrollbarTrack: "rgba(0,0,0,0.04)",
  scrollbarThumb: "rgba(0,0,0,0.25)",
  freezeLine: "#c0c0c0",
  commentIndicator: "#f59e0b",
  commentIndicatorResolved: "#9ca3af",
  remotePresenceDefault: "#4c8bf5"
};

const GRID_THEME_KEYS: ReadonlyArray<keyof GridTheme> = [
  "gridBg",
  "gridLine",
  "headerBg",
  "headerText",
  "cellText",
  "errorText",
  "selectionFill",
  "selectionBorder",
  "selectionHandle",
  "scrollbarTrack",
  "scrollbarThumb",
  "freezeLine",
  "commentIndicator",
  "commentIndicatorResolved",
  "remotePresenceDefault"
];

function applyThemeOverrides(target: GridTheme, source: Partial<GridTheme>): void {
  for (const key of GRID_THEME_KEYS) {
    const value = source[key];
    if (typeof value !== "string") continue;
    const trimmed = value.trim();
    if (!trimmed) continue;
    target[key] = trimmed;
  }
}

export function resolveGridTheme(...overrides: Array<Partial<GridTheme> | null | undefined>): GridTheme {
  const theme: GridTheme = { ...DEFAULT_GRID_THEME };
  for (const override of overrides) {
    if (!override) continue;
    applyThemeOverrides(theme, override);
  }
  return theme;
}

export function gridThemesEqual(a: GridTheme, b: GridTheme): boolean {
  for (const key of GRID_THEME_KEYS) {
    if (a[key] !== b[key]) return false;
  }
  return true;
}
