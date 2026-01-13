export const THEME_PREFERENCES: readonly ["system", "light", "dark", "high-contrast"];

export const DEFAULT_THEME_PREFERENCE: "light";

export function isThemePreference(value: any): value is (typeof THEME_PREFERENCES)[number];

export function createMemoryStorage(): {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
};

export function getDefaultStorage(): { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem?(key: string): void };

export function loadAppearanceSettings(storage?: { getItem(key: string): string | null }): { themePreference: (typeof THEME_PREFERENCES)[number] };

export function saveAppearanceSettings(
  settings: { themePreference?: any },
  storage?: { setItem(key: string, value: string): void },
): void;

export function getThemePreference(storage?: { getItem(key: string): string | null }): (typeof THEME_PREFERENCES)[number];

export function setThemePreference(
  themePreference: any,
  storage?: { setItem(key: string, value: string): void },
): void;
