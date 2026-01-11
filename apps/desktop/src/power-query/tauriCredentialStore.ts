import { credentialScopeKey } from "../../../../packages/power-query/src/credentials/store.js";

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

type CredentialEntry = { id: string; secret: unknown };

/**
 * Desktop credential store backed by Tauri commands.
 *
 * This persists secrets using an OS-keychain-backed encryption key + an encrypted
 * blob on disk (see `apps/desktop/src-tauri/src/storage/power_query_credentials.rs`).
 */
export class TauriCredentialStore {
  private invoke: TauriInvoke;

  constructor(opts?: { invoke?: TauriInvoke }) {
    this.invoke = opts?.invoke ?? getTauriInvoke();
  }

  async get(scope: any): Promise<CredentialEntry | null> {
    const scopeKey = credentialScopeKey(scope);
    const payload = await this.invoke("power_query_credential_get", { scope_key: scopeKey });
    if (payload == null) return null;
    if (!payload || typeof payload !== "object") return null;
    const entry = payload as any;
    if (typeof entry.id !== "string" || entry.id.length === 0) return null;
    return { id: entry.id, secret: entry.secret };
  }

  async set(scope: any, secret: unknown): Promise<CredentialEntry> {
    const scopeKey = credentialScopeKey(scope);
    const payload = await this.invoke("power_query_credential_set", { scope_key: scopeKey, secret });
    if (!payload || typeof payload !== "object") {
      throw new Error("Unexpected credential payload returned from Tauri");
    }
    const entry = payload as any;
    if (typeof entry.id !== "string" || entry.id.length === 0) {
      throw new Error("Credential payload missing id");
    }
    return { id: entry.id, secret: entry.secret };
  }

  async delete(scope: any): Promise<void> {
    const scopeKey = credentialScopeKey(scope);
    await this.invoke("power_query_credential_delete", { scope_key: scopeKey });
  }

  /**
   * Optional debugging helper; not part of the core `CredentialStore` contract.
   */
  async list(): Promise<Array<{ scopeKey: string; id: string }>> {
    const payload = await this.invoke("power_query_credential_list");
    if (!Array.isArray(payload)) return [];
    return payload
      .filter((e) => e && typeof e === "object")
      .map((e: any) => ({ scopeKey: String(e.scopeKey ?? ""), id: String(e.id ?? "") }))
      .filter((e) => e.scopeKey.length > 0 && e.id.length > 0);
  }
}
