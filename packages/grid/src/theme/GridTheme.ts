export interface GridTheme {
  gridBg: string;
  gridLine: string;
  headerBg: string;
  headerText: string;
  cellText: string;
  errorText: string;
  selectionFill: string;
  selectionBorder: string;
  selectionHandle: string;
  scrollbarTrack: string;
  scrollbarThumb: string;
  freezeLine: string;
  commentIndicator: string;
  commentIndicatorResolved: string;
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

