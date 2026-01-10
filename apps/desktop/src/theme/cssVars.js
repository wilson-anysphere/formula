/**
 * Resolve a CSS custom property from a root element, following simple `var(--x)`
 * references used by derived tokens.
 *
 * Note: `getComputedStyle(...).getPropertyValue("--token")` returns the *specified*
 * value for custom properties (e.g. it may return `var(--border)`), so we resolve
 * a small subset of `var()` indirections here for canvas renderers.
 *
 * @param {string} varName e.g. "--bg-primary"
 * @param {{ root?: HTMLElement | null, fallback?: string, maxDepth?: number }} [options]
 */
export function resolveCssVar(
  varName,
  { root = globalThis?.document?.documentElement ?? null, fallback = "", maxDepth = 8 } = {},
) {
  if (!root || typeof getComputedStyle !== "function") return fallback;

  let current = varName;
  const seen = new Set();

  for (let depth = 0; depth < maxDepth; depth += 1) {
    if (seen.has(current)) break;
    seen.add(current);

    const raw = getComputedStyle(root).getPropertyValue(current);
    const value = typeof raw === "string" ? raw.trim() : "";
    if (!value) return fallback;

    const varMatch = /^var\(\s*(--[^,\s)]+)\s*(?:,\s*([^)]+))?\)$/.exec(value);
    if (!varMatch) return value;

    current = varMatch[1];
    if (!current) break;
  }

  return fallback;
}

