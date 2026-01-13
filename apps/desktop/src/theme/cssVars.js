/**
 * Resolve a CSS custom property from a root element, following simple `var(--x)`
 * references used by derived tokens (including `var(--x, fallback)`).
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

  const read = (name) => {
    const raw = getComputedStyle(root).getPropertyValue(name);
    return typeof raw === "string" ? raw.trim() : "";
  };

  const parseVar = (value) => {
    const trimmed = String(value ?? "").trim();
    if (!trimmed.startsWith("var(") || !trimmed.endsWith(")")) return null;

    const inner = trimmed.slice(4, -1).trim();
    if (!inner.startsWith("--")) return null;

    let depth = 0;
    let comma = -1;
    let inSingle = false;
    let inDouble = false;

    for (let i = 0; i < inner.length; i += 1) {
      const ch = inner[i];
      const prev = i > 0 ? inner[i - 1] : "";

      if (inSingle) {
        if (ch === "'" && prev !== "\\") inSingle = false;
        continue;
      }
      if (inDouble) {
        if (ch === '"' && prev !== "\\") inDouble = false;
        continue;
      }

      if (ch === "'") {
        inSingle = true;
        continue;
      }
      if (ch === '"') {
        inDouble = true;
        continue;
      }

      if (ch === "(") depth += 1;
      else if (ch === ")") depth = Math.max(0, depth - 1);
      else if (ch === "," && depth === 0) {
        comma = i;
        break;
      }
    }

    const name = comma === -1 ? inner : inner.slice(0, comma);
    const varName = name.trim();
    if (!varName.startsWith("--")) return null;

    const fallbackValue = comma === -1 ? null : inner.slice(comma + 1).trim() || null;
    return { name: varName, fallback: fallbackValue };
  };

  const resolveValue = (value, seen) => {
    let current = String(value ?? "").trim();
    let lastFallback = null;

    for (let depth = 0; depth < maxDepth; depth += 1) {
      const parsed = parseVar(current);
      if (!parsed) return current || fallback;

      const nextName = parsed.name;
      if (parsed.fallback != null) lastFallback = parsed.fallback;

      // Handle cycles and enforce a max indirection depth.
      if (seen.has(nextName)) {
        current = parsed.fallback ?? lastFallback ?? "";
        continue;
      }
      seen.add(nextName);

      const nextValue = read(nextName);
      if (nextValue) {
        current = nextValue;
        continue;
      }

      const fb = parsed.fallback ?? lastFallback;
      if (fb != null) return resolveValue(fb, seen);
      return fallback;
    }

    if (lastFallback != null) return resolveValue(lastFallback, seen);
    return fallback;
  };

  const start = read(varName);
  if (!start) return fallback;
  return resolveValue(start, new Set([varName]));
}
