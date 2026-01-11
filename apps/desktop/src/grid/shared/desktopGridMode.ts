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

export function resolveDesktopGridMode(search: string = typeof window !== "undefined" ? window.location.search : ""): DesktopGridMode {
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

  const env = readEnvFlag();
  if (env) {
    const normalized = env.trim().toLowerCase();
    if (normalized === "shared" || normalized === "1" || normalized === "true") return "shared";
    if (normalized === "legacy" || normalized === "0" || normalized === "false") return "legacy";
  }

  return "legacy";
}

