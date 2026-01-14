import { credentialScopeKey } from "@formula/power-query";

import { getTauriInvokeOrThrow, hasTauriInvoke as hasTauriInvokeRuntime, type TauriInvoke } from "../tauri/api";

export function hasTauriInvoke(): boolean {
  return hasTauriInvokeRuntime();
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
    this.invoke = opts?.invoke ?? getTauriInvokeOrThrow();
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
