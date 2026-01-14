export type DesktopGridMode = "legacy" | "shared";

function readEnvFlag(): string | null {
  // Vite exposes env vars via `import.meta.env`, but tests may also set Node-style env.
  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  const viteValue = metaEnv?.VITE_DESKTOP_GRID_MODE ?? metaEnv?.VITE_GRID_MODE ?? metaEnv?.VITE_USE_SHARED_GRID;
  if (typeof viteValue === "string") return viteValue;
  if (typeof viteValue === "boolean") return viteValue ? "shared" : "legacy";

  const nodeEnv = (globalThis as any)?.process?.env as Record<string, unknown> | undefined;
  const nodeValue = nodeEnv?.DESKTOP_GRID_MODE ?? nodeEnv?.GRID_MODE ?? nodeEnv?.USE_SHARED_GRID;
  if (typeof nodeValue === "string") return nodeValue;
  if (typeof nodeValue === "boolean") return nodeValue ? "shared" : "legacy";
  return null;
}

/**
 * Resolve the desktop grid renderer mode.
 *
 * Precedence:
 * 1) Query param overrides (`?grid=legacy|shared`)
 * 2) Environment overrides (Vite `import.meta.env.*` or Node `process.env.*`)
 * 3) Default: `shared`
 *
 * Note: `envOverride` exists to make unit tests deterministic without needing to
 * mutate `import.meta.env` (which may be read-only depending on the runner).
 */
export function resolveDesktopGridMode(
  search: string = typeof window !== "undefined" ? window.location.search : "",
  envOverride?: string | boolean | null
): DesktopGridMode {
  try {
    const params = new URLSearchParams(search);
    const raw = params.get("grid") ?? params.get("gridMode") ?? params.get("renderer");
    if (raw) {
      const normalized = raw.trim().toLowerCase();
      if (normalized === "shared" || normalized === "new") return "shared";
      if (normalized === "legacy" || normalized === "old") return "legacy";
    }
  } catch {
    // Ignore invalid URLSearchParams input.
  }

  const env = envOverride !== undefined ? envOverride : readEnvFlag();
  if (typeof env === "boolean") return env ? "shared" : "legacy";
  if (typeof env === "string" && env.trim() !== "") {
    const normalized = env.trim().toLowerCase();
    if (normalized === "shared" || normalized === "new" || normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") {
      return "shared";
    }
    if (normalized === "legacy" || normalized === "old" || normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") {
      return "legacy";
    }
  }

  return "shared";
}
