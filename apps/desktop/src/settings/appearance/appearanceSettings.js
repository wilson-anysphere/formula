import { getDefaultStorage } from "./storage.js";

const STORAGE_KEY = "formula.settings.appearance.v1";
export const DEFAULT_THEME_PREFERENCE = "light";

/** @type {const} */
export const THEME_PREFERENCES = ["system", "light", "dark", "high-contrast"];

export function isThemePreference(value) {
  return THEME_PREFERENCES.includes(value);
}

export function loadAppearanceSettings(storage = getDefaultStorage()) {
  let raw = null;
  try {
    raw = storage?.getItem?.(STORAGE_KEY);
  } catch {
    // Storage may be unavailable (e.g. disabled in the host webview).
    // Fall back to the UX default.
    return { themePreference: DEFAULT_THEME_PREFERENCE };
  }

  if (!raw) return { themePreference: DEFAULT_THEME_PREFERENCE };

  try {
    const parsed = JSON.parse(raw);
    const themePreference = isThemePreference(parsed?.themePreference)
      ? parsed.themePreference
      : DEFAULT_THEME_PREFERENCE;

    return { themePreference };
  } catch {
    return { themePreference: DEFAULT_THEME_PREFERENCE };
  }
}

export function saveAppearanceSettings(settings, storage = getDefaultStorage()) {
  const themePreference = isThemePreference(settings?.themePreference)
    ? settings.themePreference
    : DEFAULT_THEME_PREFERENCE;

  storage.setItem(STORAGE_KEY, JSON.stringify({ themePreference }));
}

export function getThemePreference(storage = getDefaultStorage()) {
  return loadAppearanceSettings(storage).themePreference;
}

export function setThemePreference(themePreference, storage = getDefaultStorage()) {
  saveAppearanceSettings({ themePreference }, storage);
}
