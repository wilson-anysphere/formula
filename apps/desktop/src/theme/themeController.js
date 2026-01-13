import {
  getSystemReducedMotion,
  getSystemTheme,
  MEDIA,
  subscribeToMediaQuery,
} from "./systemPreferences.js";

import {
  getDefaultStorage,
  DEFAULT_THEME_PREFERENCE,
  isThemePreference,
  loadAppearanceSettings,
  setThemePreference as persistThemePreference,
} from "../settings/appearance/index.js";

export function resolveTheme(themePreference, env) {
  if (themePreference && themePreference !== "system") return themePreference;
  return getSystemTheme(env);
}

export function applyThemeToDocument(resolvedTheme, document) {
  if (!document?.documentElement) return;

  document.documentElement.setAttribute("data-theme", resolvedTheme);
}

export function applyReducedMotionToDocument(reducedMotion, document) {
  if (!document?.documentElement) return;
  document.documentElement.setAttribute(
    "data-reduced-motion",
    reducedMotion ? "true" : "false",
  );
}

export class ThemeController {
  /**
   * @param {{
   *   document?: Document,
   *   env?: { matchMedia?: (query: string) => any },
   *   storage?: { getItem(key: string): string | null, setItem(key: string, value: string): void, removeItem?(key: string): void }
   * }} [options]
   */
  constructor(options = {}) {
    this._document =
      options.document ||
      (typeof globalThis !== "undefined" ? globalThis.document : null);
    this._env = options.env || globalThis;
    this._storage = options.storage || getDefaultStorage();

    /** @type {"system" | "light" | "dark" | "high-contrast"} */
    this._themePreference = DEFAULT_THEME_PREFERENCE;
    /** @type {"light" | "dark" | "high-contrast"} */
    this._resolvedTheme = "light";

    this._unsubTheme = () => {};
    this._unsubReducedMotion = () => {};
  }

  start() {
    /** @type {{ themePreference?: "system" | "light" | "dark" | "high-contrast" }} */
    let settings = { themePreference: DEFAULT_THEME_PREFERENCE };
    try {
      settings = loadAppearanceSettings(this._storage);
    } catch {
      // Storage may be unavailable (e.g. disabled in the host webview). Themes
      // should still apply best-effort without persisting.
    }

    this.setThemePreference(settings.themePreference, { persist: false });
    this._applyReducedMotion();

    this._wireReducedMotionListener();
    this._wireThemeListener();
  }

  stop() {
    this._unsubTheme();
    this._unsubReducedMotion();
    this._unsubTheme = () => {};
    this._unsubReducedMotion = () => {};
  }

  getThemePreference() {
    return this._themePreference;
  }

  getResolvedTheme() {
    return this._resolvedTheme;
  }

  /**
   * @param {string} themePreference
   * @param {{ persist?: boolean }} [options]
   */
  setThemePreference(themePreference, options = {}) {
    const { persist = true } = options;

    const nextPreference = isThemePreference(themePreference) ? themePreference : DEFAULT_THEME_PREFERENCE;

    this._themePreference = nextPreference;
    this._applyTheme();

    if (persist) {
      try {
        persistThemePreference(nextPreference, this._storage);
      } catch {
        // Best-effort persistence only.
      }
    }

    this._wireThemeListener();
  }

  _applyTheme() {
    const resolved = resolveTheme(this._themePreference, this._env);
    this._resolvedTheme = resolved;
    applyThemeToDocument(resolved, this._document);
  }

  _applyReducedMotion() {
    const reduced = getSystemReducedMotion(this._env);
    applyReducedMotionToDocument(reduced, this._document);
  }

  _wireThemeListener() {
    this._unsubTheme();
    this._unsubTheme = () => {};

    if (this._themePreference !== "system") return;

    const onChange = () => this._applyTheme();

    // Any of these changing can affect which theme we resolve to.
    const unsubs = [
      subscribeToMediaQuery(this._env, MEDIA.forcedColors, onChange),
      subscribeToMediaQuery(this._env, MEDIA.prefersContrastMore, onChange),
      subscribeToMediaQuery(this._env, MEDIA.prefersDark, onChange)
    ];

    this._unsubTheme = () => unsubs.forEach((u) => u());
  }

  _wireReducedMotionListener() {
    this._unsubReducedMotion();
    this._unsubReducedMotion = subscribeToMediaQuery(
      this._env,
      MEDIA.reducedMotion,
      () => {
        this._applyReducedMotion();
      },
    );
  }
}
