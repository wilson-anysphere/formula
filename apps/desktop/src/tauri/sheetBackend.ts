import { getTauriInvokeOrThrow, type TauriInvoke } from "./api";

export class TauriSheetBackend {
  private readonly invoke: TauriInvoke;

  constructor(options: { invoke?: TauriInvoke } = {}) {
    this.invoke = options.invoke ?? getTauriInvokeOrThrow();
  }

  async deleteSheet(sheetId: string): Promise<void> {
    await this.invoke("delete_sheet", { sheet_id: sheetId });
  }
}
