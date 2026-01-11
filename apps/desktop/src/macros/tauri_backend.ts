import type {
  MacroBackend,
  MacroCellUpdate,
  MacroInfo,
  MacroRunRequest,
  MacroRunResult,
} from "./types";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
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

export class TauriMacroBackend implements MacroBackend {
  private readonly invoke: TauriInvoke;

  constructor(options: { invoke?: TauriInvoke } = {}) {
    this.invoke = options.invoke ?? getTauriInvoke();
  }

  async listMacros(workbookId: string): Promise<MacroInfo[]> {
    const macros = await this.invoke("list_macros", { workbook_id: workbookId });
    return macros as MacroInfo[];
  }

  async runMacro(request: MacroRunRequest): Promise<MacroRunResult> {
    const result = await this.invoke("run_macro", {
      workbook_id: request.workbookId,
      macro_id: request.macroId,
      permissions: request.permissions,
      timeout_ms: request.timeoutMs,
    });

    return {
      ok: Boolean(result.ok),
      output: Array.isArray(result.output) ? result.output.map(String) : [],
      updates: normalizeUpdates(result.updates),
      error: result.error ? { message: String(result.error.message ?? result.error) } : undefined,
    };
  }
}
