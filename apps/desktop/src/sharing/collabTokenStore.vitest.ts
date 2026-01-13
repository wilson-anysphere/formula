/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

function createInMemoryStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

function base64UrlEncodeJson(value: unknown): string {
  const json = JSON.stringify(value);
  const b64 =
    typeof btoa === "function"
      ? btoa(json)
      : // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).Buffer.from(json, "utf8").toString("base64");
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function makeJwtWithExp(expSeconds: number): string {
  const header = base64UrlEncodeJson({ alg: "none", typ: "JWT" });
  const payload = base64UrlEncodeJson({ exp: expSeconds });
  // signature can be empty for our decode-only tests.
  return `${header}.${payload}.`;
}

describe("collabTokenStore", () => {
  beforeEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();

    const storage = createInMemoryStorage();
    Object.defineProperty(globalThis, "sessionStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "sessionStorage", { configurable: true, value: storage });

    delete (globalThis as any).__TAURI__;
    delete (globalThis as any).__FORMULA_COLLAB_OPAQUE_TOKEN_TTL_MS;
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete (globalThis as any).__TAURI__;
    delete (globalThis as any).__FORMULA_COLLAB_OPAQUE_TOKEN_TTL_MS;
  });

  it("persists JWT tokens to the desktop secure store with expiresAtMs from exp", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2025-01-01T00:00:00Z"));

    const invoke = vi.fn(async () => null);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    // Ensure we get a fresh module instance (clears cached keychain store).
    vi.resetModules();
    const { storeCollabToken } = await import("./collabTokenStore.js");

    const nowSeconds = Math.floor(Date.now() / 1000);
    const expSeconds = nowSeconds + 60;
    const token = makeJwtWithExp(expSeconds);

    storeCollabToken({ wsUrl: "ws://example.com", docId: "doc-1", token });

    const tokenKey = "formula:collab:token:ws://example.com|doc-1";
    const raw = window.sessionStorage.getItem(tokenKey);
    expect(raw).toBeTruthy();
    const parsed = JSON.parse(raw!);
    expect(parsed).toEqual({ token, expiresAtMs: expSeconds * 1000 });

    expect(invoke).toHaveBeenCalledWith("collab_token_set", {
      token_key: tokenKey,
      entry: { token, expiresAtMs: expSeconds * 1000 },
    });
  });

  it("hydrates tokens from the desktop secure store during preload", async () => {
    const tokenKey = "formula:collab:token:ws://sync.example|doc-xyz";
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "collab_token_get") {
        return { token: "secret-token", expiresAtMs: Date.now() + 60_000 };
      }
      return null;
    });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    vi.resetModules();
    const { loadCollabToken, preloadCollabTokenFromKeychain } = await import("./collabTokenStore.js");

    await preloadCollabTokenFromKeychain({ wsUrl: "ws://sync.example", docId: "doc-xyz" });

    expect(invoke).toHaveBeenCalledWith("collab_token_get", { token_key: tokenKey });
    expect(loadCollabToken({ wsUrl: "ws://sync.example", docId: "doc-xyz" })).toBe("secret-token");
  });

  it("deletes expired tokens from the secure store on preload", async () => {
    const tokenKey = "formula:collab:token:ws://sync.example|doc-expired";
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "collab_token_get") {
        return { token: "expired", expiresAtMs: Date.now() - 1 };
      }
      return null;
    });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    vi.resetModules();
    const { loadCollabToken, preloadCollabTokenFromKeychain } = await import("./collabTokenStore.js");

    await preloadCollabTokenFromKeychain({ wsUrl: "ws://sync.example", docId: "doc-expired" });

    expect(invoke).toHaveBeenCalledWith("collab_token_get", { token_key: tokenKey });
    expect(invoke).toHaveBeenCalledWith("collab_token_delete", { token_key: tokenKey });
    expect(loadCollabToken({ wsUrl: "ws://sync.example", docId: "doc-expired" })).toBeNull();
  });

  it("applies a conservative TTL to opaque tokens", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2025-01-01T00:00:00Z"));

    (globalThis as any).__FORMULA_COLLAB_OPAQUE_TOKEN_TTL_MS = 10_000;

    const invoke = vi.fn(async () => null);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    vi.resetModules();
    const { loadCollabToken, storeCollabToken } = await import("./collabTokenStore.js");

    storeCollabToken({ wsUrl: "ws://opaque.example", docId: "doc-opaque", token: "opaque" });

    const tokenKey = "formula:collab:token:ws://opaque.example|doc-opaque";
    expect(loadCollabToken({ wsUrl: "ws://opaque.example", docId: "doc-opaque" })).toBe("opaque");

    expect(invoke).toHaveBeenCalledWith("collab_token_set", {
      token_key: tokenKey,
      entry: { token: "opaque", expiresAtMs: Date.now() + 10_000 },
    });
  });
});

