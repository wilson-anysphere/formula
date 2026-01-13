type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (typeof invoke !== "function") {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export function hasTauriInvoke(): boolean {
  return typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
}

export type CollabTokenKeychainEntry = {
  token: string;
  /**
   * Absolute epoch ms when the token should be considered expired.
   *
   * When the token is a JWT, this is derived from the `exp` claim (seconds since epoch).
   * For opaque tokens, the frontend can apply a conservative TTL.
   */
  expiresAtMs: number | null;
};

/**
 * Desktop collab token store backed by Tauri commands.
 *
 * This persists secrets using an OS-keychain-backed encryption key + an encrypted
 * blob on disk (see `apps/desktop/src-tauri/src/storage/collab_tokens.rs`).
 *
 * IMPORTANT:
 * - Token values must never be logged.
 */
export class CollabTokenKeychainStore {
  private invoke: TauriInvoke;

  constructor(opts?: { invoke?: TauriInvoke }) {
    this.invoke = opts?.invoke ?? getTauriInvoke();
  }

  async get(tokenKey: string): Promise<CollabTokenKeychainEntry | null> {
    const key = String(tokenKey ?? "").trim();
    if (!key) return null;
    const payload = await this.invoke("collab_token_get", { token_key: key });
    if (payload == null) return null;
    if (!payload || typeof payload !== "object") return null;
    const entry = payload as any;
    const token = typeof entry.token === "string" ? entry.token : "";
    if (!token) return null;
    const expiresAtMs =
      entry.expiresAtMs == null ? null : typeof entry.expiresAtMs === "number" ? entry.expiresAtMs : Number(entry.expiresAtMs);
    return { token, expiresAtMs: Number.isFinite(expiresAtMs as number) ? (expiresAtMs as number) : null };
  }

  async set(tokenKey: string, entry: CollabTokenKeychainEntry): Promise<void> {
    const key = String(tokenKey ?? "").trim();
    const token = typeof entry?.token === "string" ? entry.token : "";
    if (!key || !token) return;
    const expiresAtMs =
      entry.expiresAtMs == null ? null : Math.trunc(Number(entry.expiresAtMs));
    await this.invoke("collab_token_set", {
      token_key: key,
      entry: { token, expiresAtMs: Number.isFinite(expiresAtMs) ? expiresAtMs : null },
    });
  }

  async delete(tokenKey: string): Promise<void> {
    const key = String(tokenKey ?? "").trim();
    if (!key) return;
    await this.invoke("collab_token_delete", { token_key: key });
  }
}
