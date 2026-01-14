function readEnvFlag(): string | boolean | null {
  // Vite exposes env vars via `import.meta.env`, but tests may also set Node-style env.
  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  const viteValue = metaEnv?.VITE_ENABLE_DRAWING_INTERACTIONS;
  if (typeof viteValue === "string") return viteValue;
  if (typeof viteValue === "boolean") return viteValue;

  const nodeEnv = (globalThis as any)?.process?.env as Record<string, unknown> | undefined;
  const nodeValue = nodeEnv?.ENABLE_DRAWING_INTERACTIONS;
  if (typeof nodeValue === "string") return nodeValue;
  if (typeof nodeValue === "boolean") return nodeValue;
  return null;
}

function parseBooleanFlag(value: unknown): boolean | null {
  if (value === true) return true;
  if (value === false) return false;
  if (value == null) return null;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "") return null;
    if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") return true;
    if (normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") return false;
  }
  return null;
}

/**
 * Resolve whether interactive drawing manipulation (select/drag/resize/rotate) should be enabled.
 *
 * Precedence:
 * 1) Query param overrides (`?drawingInteractions=1`)
 * 2) Environment overrides (Vite `import.meta.env.*` or Node `process.env.*`)
 * 3) Default: `false`
 *
 * Note: `envOverride` exists to make unit tests deterministic without needing to
 * mutate `import.meta.env` (which may be read-only depending on the runner).
 */
export function resolveEnableDrawingInteractions(
  search: string = typeof window !== "undefined" ? window.location.search : "",
  envOverride?: string | boolean | null,
): boolean {
  try {
    const params = new URLSearchParams(search);
    const raw =
      params.get("drawingInteractions") ??
      params.get("drawings") ??
      params.get("enableDrawingInteractions") ??
      null;
    const parsed = parseBooleanFlag(raw);
    if (parsed != null) return parsed;
  } catch {
    // Ignore invalid URLSearchParams input.
  }

  const env = envOverride !== undefined ? envOverride : readEnvFlag();
  const parsedEnv = parseBooleanFlag(env);
  if (parsedEnv != null) return parsedEnv;

  return false;
}

