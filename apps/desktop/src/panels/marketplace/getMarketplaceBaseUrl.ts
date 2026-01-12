const STORAGE_KEY = "formula:marketplace:baseUrl";

// Production fallback. This is the hosted marketplace service used by the desktop app
// when no local dev server is present.
const DEFAULT_PRODUCTION_BASE_URL = "https://marketplace.formula.app/api";

function normalizeBaseUrl(value: string): string | null {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return null;

  // A convenience: users/tests may provide a marketplace *origin* instead of the full
  // API base path (e.g. "https://marketplace.formula.app" vs ".../api"). In that case,
  // default the path to `/api` to match the MarketplaceClient contract.
  try {
    // Absolute URL.
    if (/^https?:\/\//i.test(trimmed)) {
      const url = new URL(trimmed);
      if (!url.pathname || url.pathname === "/") {
        url.pathname = "/api";
      }
      // Normalize trailing slash (MarketplaceClient expects no trailing slash).
      return url.toString().replace(/\/$/, "");
    }
  } catch {
    // Ignore URL parsing errors and fall through to the string-based normalization.
  }

  // Relative path (same-origin).
  if (trimmed === "/") return "/api";

  return trimmed.endsWith("/") ? trimmed.slice(0, -1) : trimmed;
}

function tryReadLocalStorage(storage: Pick<Storage, "getItem"> | undefined): string | null {
  if (!storage) return null;
  try {
    return normalizeBaseUrl(storage.getItem(STORAGE_KEY) ?? "");
  } catch {
    return null;
  }
}

function readViteEnv(env: Record<string, unknown> | undefined): string | null {
  if (!env) return null;
  const raw = env.VITE_FORMULA_MARKETPLACE_BASE_URL;
  if (typeof raw !== "string") return null;
  return normalizeBaseUrl(raw);
}

function isProductionEnv(env: Record<string, unknown> | undefined): boolean {
  if (!env) return false;
  const prod = env.PROD;
  if (typeof prod === "boolean") return prod;
  const mode = env.MODE;
  if (typeof mode === "string") return mode === "production";
  return false;
}

export function getMarketplaceBaseUrl(options?: {
  /**
   * Override the storage implementation (useful for unit tests).
   *
   * Defaults to `globalThis.localStorage` when available.
   */
  storage?: Pick<Storage, "getItem"> | undefined;
  /**
   * Override the env implementation (useful for unit tests).
   *
   * Defaults to `import.meta.env` when available.
   */
  env?: Record<string, unknown> | undefined;
}): string {
  const storage = (() => {
    if (options?.storage) return options.storage;
    try {
      return (globalThis as any).localStorage as Pick<Storage, "getItem"> | undefined;
    } catch {
      return undefined;
    }
  })();

  const localValue = tryReadLocalStorage(storage);
  if (localValue) return localValue;

  const metaEnv = options?.env ?? ((import.meta as any)?.env as Record<string, unknown> | undefined);
  const envValue = readViteEnv(metaEnv);
  if (envValue) return envValue;

  // Default behaviour:
  // - dev/e2e: rely on same-origin `/api` stubs
  // - production: use the hosted marketplace API
  return isProductionEnv(metaEnv) ? DEFAULT_PRODUCTION_BASE_URL : "/api";
}
