const STORAGE_KEY = "formula:marketplace:baseUrl";

// Production fallback. This is the hosted marketplace service used by the desktop app
// when no local dev server is present.
const DEFAULT_PRODUCTION_BASE_URL = "https://marketplace.formula.app/api";

function normalizeBaseUrl(value: string): string | null {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return null;

  // A convenience: treat "/" as the default API base.
  if (trimmed === "/") return "/api";

  // A convenience: users/tests may provide a marketplace *origin* instead of the full
  // API base path (e.g. "https://marketplace.formula.app" vs ".../api"). In that case,
  // default the path to `/api` to match the MarketplaceClient contract.
  if (/^https?:\/\//i.test(trimmed)) {
    let url: URL;
    try {
      url = new URL(trimmed);
    } catch {
      // Treat invalid absolute URL overrides as unset so we fall back to safe defaults.
      return null;
    }

    // Base URL should not carry query/hash.
    url.search = "";
    url.hash = "";

    let pathname = url.pathname.replace(/\/+$/, "");
    if (!pathname || pathname === "/") pathname = "/api";
    url.pathname = pathname;

    // MarketplaceClient expects no trailing slash.
    return `${url.origin}${url.pathname}`;
  }

  // Relative path (same-origin). Keep it tolerant; MarketplaceClient will normalize further.
  let out = trimmed.replace(/\/+$/, "");
  while (out.startsWith("./")) out = out.slice(2);
  if (!out) return null;
  if (!out.startsWith("/")) out = `/${out}`;
  return out;
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
