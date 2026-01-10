import { enUS } from "./locales/en-US.js";
import { deDE } from "./locales/de-DE.js";
import { ar } from "./locales/ar.js";

const LOCALES = {
  "en-US": { direction: "ltr", messages: enUS },
  "de-DE": { direction: "ltr", messages: deDE },
  ar: { direction: "rtl", messages: ar }
};

let currentLocale = "en-US";

export function availableLocales() {
  return Object.keys(LOCALES);
}

export function getLocale() {
  return currentLocale;
}

export function getDirection(locale = currentLocale) {
  return LOCALES[locale]?.direction ?? "ltr";
}

export function setLocale(locale) {
  if (!LOCALES[locale]) {
    throw new Error(`Unsupported locale: ${locale}`);
  }
  currentLocale = locale;

  // RTL hook: reflect locale direction to the root document so CSS logical
  // properties (`margin-inline-start`, etc.) can flip layout automatically.
  if (typeof document !== "undefined" && document?.documentElement) {
    document.documentElement.lang = locale;
    document.documentElement.dir = getDirection(locale);
  }
}

export function t(key) {
  const active = LOCALES[currentLocale]?.messages ?? {};
  const fallback = LOCALES["en-US"]?.messages ?? {};
  const template = active[key] ?? fallback[key] ?? key;
  return template;
}

/**
 * Translate a key with simple `{var}` interpolation.
 *
 * @param {string} key
 * @param {Record<string, unknown>} vars
 */
export function tWithVars(key, vars) {
  const template = t(key);
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (_, name) => {
    const value = vars[name];
    return value == null ? `{${name}}` : String(value);
  });
}
