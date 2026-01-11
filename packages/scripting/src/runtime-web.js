/**
 * Browser/desktop-webview ScriptRuntime powered by a WebWorker.
 *
 * This runtime mirrors the Node `worker_threads` implementation but relies only
 * on standard Web APIs so it can run inside Vite/Tauri webviews.
 */

/**
 * @typedef {{ level: "log" | "info" | "warn" | "error", message: string }} ScriptConsoleEntry
 * @typedef {{ logs: ScriptConsoleEntry[], error?: { name?: string, message: string, stack?: string } }} ScriptRunResult
 * @typedef {import("./workbook.js").Workbook} Workbook
 *
 * @typedef {{
 *   network?: "none" | "allowlist" | "full",
 *   networkAllowlist?: string[],
 * }} ScriptPermissions
 */

const WORKER_URL = new URL("./web-sandbox-worker.js", import.meta.url);

export class ScriptRuntime {
  /**
   * @param {Workbook} workbook
   */
  constructor(workbook) {
    this.workbook = workbook;
  }

  /**
   * @param {string} code
   * @param {{ permissions?: ScriptPermissions }=} options
   * @returns {Promise<ScriptRunResult>}
   */
  async run(code, options) {
    if (typeof Worker === "undefined") {
      throw new Error("ScriptRuntime requires Web Worker support");
    }

    const activeSheetName = this.workbook.getActiveSheet().name;
    const selection = this.workbook.getSelection();

    /** @type {ScriptConsoleEntry[]} */
    const logs = [];

    const worker = new Worker(WORKER_URL, { type: "module" });

    const completion = new Promise((resolve) => {
      const cleanup = () => {
        worker.onmessage = null;
        worker.onerror = null;
        worker.terminate();
      };

      worker.onmessage = async (event) => {
        const message = event.data;

        if (message?.type === "console") {
          logs.push({ level: message.level, message: message.message });
          return;
        }

        if (message?.type === "rpc") {
          try {
            const result = await this.handleRpc(message.method, message.params);
            worker.postMessage({ type: "rpcResult", id: message.id, result });
          } catch (err) {
            worker.postMessage({ type: "rpcError", id: message.id, error: serializeError(err) });
          }
          return;
        }

        if (message?.type === "result") {
          cleanup();
          resolve({ logs });
          return;
        }

        if (message?.type === "error") {
          cleanup();
          resolve({ logs, error: message.error });
        }
      };

      worker.onerror = (event) => {
        cleanup();
        resolve({ logs, error: serializeError(event?.error ?? event?.message ?? "Worker error") });
      };
    });

    worker.postMessage({
      type: "run",
      code,
      activeSheetName,
      selection,
      permissions: options?.permissions,
    });

    return completion;
  }

  async handleRpc(method, params) {
    switch (method) {
      case "range.getValues": {
        const { sheetName, address } = params;
        return this.workbook.getSheet(sheetName).getRange(address).getValues();
      }
      case "range.setValues": {
        const { sheetName, address, values } = params;
        this.workbook.getSheet(sheetName).getRange(address).setValues(values);
        return null;
      }
      case "range.getValue": {
        const { sheetName, address } = params;
        return this.workbook.getSheet(sheetName).getRange(address).getValue();
      }
      case "range.setValue": {
        const { sheetName, address, value } = params;
        this.workbook.getSheet(sheetName).getRange(address).setValue(value);
        return null;
      }
      case "range.getFormat": {
        const { sheetName, address } = params;
        return this.workbook.getSheet(sheetName).getRange(address).getFormat();
      }
      case "range.setFormat": {
        const { sheetName, address, format } = params;
        this.workbook.getSheet(sheetName).getRange(address).setFormat(format);
        return null;
      }
      case "workbook.getActiveSheetName": {
        return this.workbook.getActiveSheet().name;
      }
      case "workbook.getSelection": {
        return this.workbook.getSelection();
      }
      case "workbook.setSelection": {
        const { sheetName, address } = params;
        this.workbook.setSelection(sheetName, address);
        return null;
      }
      default:
        throw new Error(`Unknown RPC method: ${method}`);
    }
  }
}

function serializeError(err) {
  if (err instanceof Error) {
    return { message: err.message, name: err.name, stack: err.stack };
  }
  if (typeof err === "string") {
    return { message: err };
  }
  return { message: "Unknown error" };
}
