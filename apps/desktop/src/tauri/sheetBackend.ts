type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export class TauriSheetBackend {
  private readonly invoke: TauriInvoke;

  constructor(options: { invoke?: TauriInvoke } = {}) {
    this.invoke = options.invoke ?? getTauriInvoke();
  }

  async deleteSheet(sheetId: string): Promise<void> {
    await this.invoke("delete_sheet", { sheet_id: sheetId });
  }
}

