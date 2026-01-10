export const MEDIA = {
  prefersDark: "(prefers-color-scheme: dark)",
  forcedColors: "(forced-colors: active)",
  prefersContrastMore: "(prefers-contrast: more)",
  reducedMotion: "(prefers-reduced-motion: reduce)"
};

function getMatchMedia(env) {
  return env && typeof env.matchMedia === "function" ? env.matchMedia.bind(env) : null;
}

export function getSystemTheme(env = globalThis) {
  const matchMedia = getMatchMedia(env);
  if (!matchMedia) return "light";

  if (
    matchMedia(MEDIA.forcedColors).matches ||
    matchMedia(MEDIA.prefersContrastMore).matches
  ) {
    return "high-contrast";
  }

  return matchMedia(MEDIA.prefersDark).matches ? "dark" : "light";
}

export function getSystemReducedMotion(env = globalThis) {
  const matchMedia = getMatchMedia(env);
  if (!matchMedia) return false;
  return matchMedia(MEDIA.reducedMotion).matches;
}

export function subscribeToMediaQuery(env, query, onChange) {
  const matchMedia = getMatchMedia(env);
  if (!matchMedia) return () => {};

  const mql = matchMedia(query);
  const handler = (event) => onChange(Boolean(event?.matches));

  if (typeof mql.addEventListener === "function") {
    mql.addEventListener("change", handler);
    return () => mql.removeEventListener("change", handler);
  }

  // Older webviews / Safari.
  if (typeof mql.addListener === "function") {
    mql.addListener(handler);
    return () => mql.removeListener(handler);
  }

  return () => {};
}
