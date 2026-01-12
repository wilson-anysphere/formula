import type {
  MacroBackend,
  MacroBlockedError,
  MacroCellUpdate,
  MacroInfo,
  MacroPermissionRequest,
  MacroRunRequest,
  MacroRunResult,
  MacroSecurityStatus,
  MacroSignatureInfo,
  MacroSignatureStatus,
  MacroTrustDecision,
} from "./types";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

function nonNegativeInt(value: unknown): number {
  const num = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(num)) return 0;
  const floored = Math.floor(num);
  if (!Number.isSafeInteger(floored) || floored < 0) return 0;
  return floored;
}

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

function errorMessage(err: unknown): string {
  if (typeof err === "string") return err;
  if (err instanceof Error) return err.message;
  if (err && typeof err === "object" && "message" in err) {
    try {
      return String((err as any).message);
    } catch {
      return "Unknown error";
    }
  }
  try {
    return String(err);
  } catch {
    return "Unknown error";
  }
}

function isNoWorkbookLoadedError(err: unknown): boolean {
  return errorMessage(err).toLowerCase().includes("no workbook loaded");
}

function normalizeUpdates(raw: any[] | undefined): MacroCellUpdate[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0) return undefined;
  const out: MacroCellUpdate[] = [];
  for (const u of raw) {
    if (!u || typeof u !== "object") continue;
    const sheetId = String((u as any).sheet_id ?? "").trim();
    const row = Number((u as any).row);
    const col = Number((u as any).col);
    if (!sheetId) continue;
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    out.push({
      sheetId,
      row,
      col,
      value: (u as any).value ?? null,
      formula: typeof (u as any).formula === "string" ? (u as any).formula : null,
      displayValue: String((u as any).display_value ?? ""),
    });
  }
  return out.length > 0 ? out : undefined;
}

function normalizeSignatureInfo(raw: any): MacroSignatureInfo | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const statusRaw = (raw as any).status;
  const status = typeof statusRaw === "string" ? (statusRaw as MacroSignatureStatus) : "unsigned";
  const signerSubject =
    typeof (raw as any).signer_subject === "string" ? String((raw as any).signer_subject) : undefined;
  const signatureBase64 =
    typeof (raw as any).signature_base64 === "string" ? String((raw as any).signature_base64) : undefined;
  return {
    status,
    signerSubject,
    signatureBase64,
  };
}

function normalizeMacroSecurityStatus(raw: any): MacroSecurityStatus {
  const hasMacros = Boolean(raw?.has_macros);
  const originPath = typeof raw?.origin_path === "string" ? String(raw.origin_path) : undefined;
  const workbookFingerprint =
    typeof raw?.workbook_fingerprint === "string" ? String(raw.workbook_fingerprint) : undefined;
  const signature = normalizeSignatureInfo(raw?.signature);
  const trustRaw = raw?.trust;
  const trust = typeof trustRaw === "string" ? (trustRaw as MacroTrustDecision) : "blocked";
  return {
    hasMacros,
    originPath,
    workbookFingerprint,
    signature,
    trust,
  };
}

function normalizeMacroBlockedError(raw: any): MacroBlockedError | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const reason = typeof (raw as any).reason === "string" ? String((raw as any).reason) : undefined;
  const status = (raw as any).status;
  if (!reason || !status) return undefined;
  return {
    reason: reason as MacroBlockedError["reason"],
    status: normalizeMacroSecurityStatus(status),
  };
}

function normalizePermissionRequest(raw: any): MacroPermissionRequest | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const reason = typeof (raw as any).reason === "string" ? String((raw as any).reason) : "";
  const macroId = typeof (raw as any).macro_id === "string" ? String((raw as any).macro_id) : "";
  if (!reason || !macroId) return undefined;
  const workbookOriginPath =
    typeof (raw as any).workbook_origin_path === "string" ? String((raw as any).workbook_origin_path) : undefined;
  const requestedRaw = (raw as any).requested;
  const requested = Array.isArray(requestedRaw) ? requestedRaw.map(String) : [];
  return { reason, macroId, workbookOriginPath, requested: requested as MacroPermissionRequest["requested"] };
}
export type MacroSelectionRect = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export type MacroUiContext = {
  sheetId: string;
  activeRow: number;
  activeCol: number;
  selection?: MacroSelectionRect | null;
};

export class TauriMacroBackend implements MacroBackend {
  private readonly invoke: TauriInvoke;

  constructor(options: { invoke?: TauriInvoke } = {}) {
    this.invoke = options.invoke ?? getTauriInvoke();
  }

  async listMacros(workbookId: string): Promise<MacroInfo[]> {
    try {
      const macros = await this.invoke("list_macros", { workbook_id: workbookId });
      return macros as MacroInfo[];
    } catch (err) {
      if (isNoWorkbookLoadedError(err)) return [];
      throw err;
    }
  }

  async getMacroSecurityStatus(workbookId: string): Promise<MacroSecurityStatus> {
    try {
      const status = await this.invoke("get_macro_security_status", { workbook_id: workbookId });
      return normalizeMacroSecurityStatus(status);
    } catch (err) {
      if (isNoWorkbookLoadedError(err)) return { hasMacros: false, trust: "blocked" };
      throw err;
    }
  }

  async setMacroTrust(workbookId: string, decision: MacroTrustDecision): Promise<MacroSecurityStatus> {
    try {
      const status = await this.invoke("set_macro_trust", { workbook_id: workbookId, decision });
      return normalizeMacroSecurityStatus(status);
    } catch (err) {
      if (isNoWorkbookLoadedError(err)) return { hasMacros: false, trust: "blocked" };
      throw err;
    }
  }

  async runMacro(request: MacroRunRequest): Promise<MacroRunResult> {
    let result: any;
    try {
      result = await this.invoke("run_macro", {
        workbook_id: request.workbookId,
        macro_id: request.macroId,
        permissions: request.permissions,
        timeout_ms: request.timeoutMs,
      });
    } catch (err) {
      if (isNoWorkbookLoadedError(err)) {
        return { ok: false, output: [], error: { message: "No workbook loaded." } };
      }
      throw err;
    }

    return {
      ok: Boolean(result.ok),
      output: Array.isArray(result.output) ? result.output.map(String) : [],
      updates: normalizeUpdates(result.updates),
      permissionRequest: normalizePermissionRequest(result.permission_request),
      error: result.error
        ? {
            message: String(result.error.message ?? result.error),
            stack: typeof result.error.stack === "string" ? String(result.error.stack) : undefined,
            code: typeof result.error.code === "string" ? String(result.error.code) : undefined,
            blocked: normalizeMacroBlockedError(result.error.blocked),
          }
        : undefined,
    };
  }

  async setMacroUiContext(options: { workbookId: string } & MacroUiContext): Promise<void> {
    const selection = (() => {
      if (!options.selection) return null;
      const startRow = nonNegativeInt(options.selection.startRow);
      const endRow = nonNegativeInt(options.selection.endRow);
      const startCol = nonNegativeInt(options.selection.startCol);
      const endCol = nonNegativeInt(options.selection.endCol);
      return {
        start_row: Math.min(startRow, endRow),
        start_col: Math.min(startCol, endCol),
        end_row: Math.max(startRow, endRow),
        end_col: Math.max(startCol, endCol),
      };
    })();

    try {
      await this.invoke("set_macro_ui_context", {
        workbook_id: options.workbookId,
        sheet_id: options.sheetId,
        active_row: nonNegativeInt(options.activeRow),
        active_col: nonNegativeInt(options.activeCol),
        selection,
      });
    } catch (err) {
      if (isNoWorkbookLoadedError(err)) return;
      // Older backends may not implement UI context sync; macro runs should still work.
      console.warn("Failed to sync macro UI context:", err);
    }
  }
}

export function wrapTauriMacroBackendWithUiContext(
  backend: TauriMacroBackend,
  getContext: () => MacroUiContext,
  options: { beforeRunMacro?: () => Promise<void> } = {}
): MacroBackend {
  return {
    listMacros: (workbookId) => backend.listMacros(workbookId),
    getMacroSecurityStatus: (workbookId) => backend.getMacroSecurityStatus(workbookId),
    setMacroTrust: (workbookId, decision) => backend.setMacroTrust(workbookId, decision),
    runMacro: async (request) => {
      if (options.beforeRunMacro) {
        await options.beforeRunMacro();
      }

      const ctx = getContext();
      await backend.setMacroUiContext({ workbookId: request.workbookId, ...ctx });
      return backend.runMacro(request);
    },
  };
}
