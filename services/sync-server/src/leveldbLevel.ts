import { createRequire } from "node:module";

function normalizeLevelExport(mod: unknown): unknown {
  if (typeof mod === "function") return mod;
  if (
    typeof mod === "object" &&
    mod !== null &&
    "default" in mod &&
    typeof (mod as { default?: unknown }).default === "function"
  ) {
    return (mod as { default: unknown }).default;
  }
  return null;
}

/**
 * Resolve the `level` module in a way that works with both npm/yarn (hoisted deps)
 * and pnpm (strict node_modules with no transitive hoisting).
 *
 * When `y-leveldb` is installed, it always has `level` as a dependency. Under pnpm,
 * that dependency is only resolvable from within the `y-leveldb` package boundary,
 * not from the sync-server package itself.
 */
export function requireLevelForYLeveldb(): (location: string, opts: any) => any {
  const require = createRequire(import.meta.url);

  // First try the simple case: `level` is available as a direct/hoisted dependency.
  try {
    const mod = require("level");
    const normalized = normalizeLevelExport(mod);
    if (normalized) return normalized as (location: string, opts: any) => any;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code !== "MODULE_NOT_FOUND") throw err;
  }

  // pnpm case: resolve `level` relative to `y-leveldb`.
  try {
    const yLeveldbRequire = createRequire(require.resolve("y-leveldb"));
    const mod = yLeveldbRequire("level");
    const normalized = normalizeLevelExport(mod);
    if (normalized) return normalized as (location: string, opts: any) => any;
    throw new Error("Unexpected 'level' export shape");
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Failed to resolve the 'level' module required for encrypted LevelDB persistence (${reason}). Ensure 'y-leveldb' is installed with its dependencies.`
    );
  }
}

