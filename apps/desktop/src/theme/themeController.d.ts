export type ThemePreference = "system" | "light" | "dark" | "high-contrast";
export type ResolvedTheme = "light" | "dark" | "high-contrast";

export function resolveTheme(themePreference: ThemePreference | null | undefined, env: any): ResolvedTheme;

export function applyThemeToDocument(resolvedTheme: ResolvedTheme, document: Document): void;

export function applyReducedMotionToDocument(reducedMotion: boolean, document: Document): void;

export class ThemeController {
  constructor(options?: {
    document?: Document;
    env?: { matchMedia?: (query: string) => any };
    storage?: { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem?(key: string): void };
  });

  start(): void;
  stop(): void;

  getThemePreference(): ThemePreference;
  getResolvedTheme(): ResolvedTheme;

  setThemePreference(themePreference: string, options?: { persist?: boolean }): void;
}

