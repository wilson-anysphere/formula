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
  if (!Array.isArray(raw)) return undefined;
  return raw.map((u) => ({
    sheetId: String(u.sheet_id ?? ""),
    row: Number(u.row ?? 0),
    col: Number(u.col ?? 0),
    value: u.value ?? null,
    formula: u.formula ?? null,
    displayValue: String(u.display_value ?? ""),
  }));
}

export class TauriMacroBackend implements MacroBackend {
  private readonly invoke: TauriInvoke;

  constructor() {
    this.invoke = getTauriInvoke();
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

