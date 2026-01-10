import { getDefaultStorage } from "./storage.js";

const STORAGE_KEY = "formula.settings.appearance.v1";

/** @type {const} */
export const THEME_PREFERENCES = ["system", "light", "dark", "high-contrast"];

export function isThemePreference(value) {
  return THEME_PREFERENCES.includes(value);
}

export function loadAppearanceSettings(storage = getDefaultStorage()) {
  const raw = storage.getItem(STORAGE_KEY);
  if (!raw) {
    return { themePreference: "system" };
  }

  try {
    const parsed = JSON.parse(raw);
    const themePreference = isThemePreference(parsed?.themePreference)
      ? parsed.themePreference
      : "system";

    return { themePreference };
  } catch {
    return { themePreference: "system" };
  }
}

export function saveAppearanceSettings(settings, storage = getDefaultStorage()) {
  const themePreference = isThemePreference(settings?.themePreference)
    ? settings.themePreference
    : "system";

  storage.setItem(STORAGE_KEY, JSON.stringify({ themePreference }));
}

export function getThemePreference(storage = getDefaultStorage()) {
  return loadAppearanceSettings(storage).themePreference;
}

export function setThemePreference(themePreference, storage = getDefaultStorage()) {
  saveAppearanceSettings({ themePreference }, storage);
}
