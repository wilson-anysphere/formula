type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void };

function safeStorage(storage: StorageLike): StorageLike {
  return {
    getItem(key) {
      try {
        return storage.getItem(key);
      } catch {
        return null;
      }
    },
    setItem(key, value) {
      try {
        storage.setItem(key, value);
      } catch {
        // ignore
      }
    },
    removeItem(key) {
      try {
        storage.removeItem(key);
      } catch {
        // ignore
      }
    },
  };
}

function getLocalStorageOrNull(): StorageLike | null {
  try {
    const storage = (globalThis as any)?.localStorage as StorageLike | undefined;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
      return safeStorage(storage);
    }
  } catch {
    // ignore
  }
  return null;
}

function storageKey(workbookId: string | undefined): string {
  return `formula.desktop.powerQuery.oauthProviders:${workbookId ?? "default"}`;
}

function safeParseJson(text: string): unknown | null {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

export type OAuth2ProviderConfig = {
  id: string;
  clientId: string;
  clientSecret?: string;
  tokenEndpoint: string;
  authorizationEndpoint?: string;
  redirectUri?: string;
  deviceAuthorizationEndpoint?: string;
  defaultScopes?: string[];
  authorizationParams?: Record<string, string>;
};

export function loadOAuth2ProviderConfigs(
  workbookId: string | undefined,
  opts: { storage?: StorageLike | null } = {},
): OAuth2ProviderConfig[] {
  const storage = opts.storage ?? getLocalStorageOrNull();
  if (!storage) return [];
  const raw = storage.getItem(storageKey(workbookId));
  if (!raw) return [];
  const parsed = safeParseJson(raw);
  if (!Array.isArray(parsed)) return [];
  return parsed
    .filter((p: any) => p && typeof p === "object" && typeof p.id === "string" && typeof p.clientId === "string" && typeof p.tokenEndpoint === "string")
    .map((p: any) => ({
      id: String(p.id),
      clientId: String(p.clientId),
      tokenEndpoint: String(p.tokenEndpoint),
      ...(typeof p.clientSecret === "string" && p.clientSecret ? { clientSecret: String(p.clientSecret) } : {}),
      ...(typeof p.authorizationEndpoint === "string" && p.authorizationEndpoint ? { authorizationEndpoint: String(p.authorizationEndpoint) } : {}),
      ...(typeof p.redirectUri === "string" && p.redirectUri ? { redirectUri: String(p.redirectUri) } : {}),
      ...(typeof p.deviceAuthorizationEndpoint === "string" && p.deviceAuthorizationEndpoint ? { deviceAuthorizationEndpoint: String(p.deviceAuthorizationEndpoint) } : {}),
      ...(Array.isArray(p.defaultScopes) ? { defaultScopes: p.defaultScopes.filter((s: any) => typeof s === "string") } : {}),
      ...(p.authorizationParams && typeof p.authorizationParams === "object" ? { authorizationParams: p.authorizationParams } : {}),
    }));
}

export function saveOAuth2ProviderConfigs(
  workbookId: string | undefined,
  configs: OAuth2ProviderConfig[],
  opts: { storage?: StorageLike | null } = {},
): void {
  const storage = opts.storage ?? getLocalStorageOrNull();
  if (!storage) return;
  storage.setItem(storageKey(workbookId), JSON.stringify(configs));
}

