import { Worker } from "node:worker_threads";

/**
 * @typedef {{ level: "log" | "info" | "warn" | "error", message: string }} ScriptConsoleEntry
 * @typedef {{ logs: ScriptConsoleEntry[], error?: { name?: string, message: string, stack?: string } }} ScriptRunResult
 * @typedef {import("./workbook.js").Workbook} Workbook
 */

const WORKER_URL = new URL("./sandbox-worker.cjs", import.meta.url);
const DEFAULT_TIMEOUT_MS = 5_000;

export class ScriptRuntime {
  /**
   * @param {Workbook} workbook
   */
  constructor(workbook) {
    this.workbook = workbook;
  }

  /**
   * @param {string} code
   * @param {{ permissions?: any, timeoutMs?: number }=} options
   * @returns {Promise<ScriptRunResult>}
   *
   * If execution exceeds `timeoutMs`, the worker is terminated and the promise
   * resolves with a `ScriptRunResult` containing an error.
   */
  async run(code, options) {
    const activeSheetName = this.workbook.getActiveSheet().name;
    const selection = this.workbook.getSelection();

    /** @type {ScriptConsoleEntry[]} */
    const logs = [];

    const worker = new Worker(WORKER_URL, {
      workerData: {
        activeSheetName,
        selection,
        permissions: options?.permissions,
      },
    });

    const timeoutMs =
      Number.isFinite(options?.timeoutMs) && options.timeoutMs > 0 ? options.timeoutMs : DEFAULT_TIMEOUT_MS;
    let timeoutId;
    let settled = false;

    const completion = new Promise((resolve) => {
      const settle = async (result) => {
        if (settled) return;
        settled = true;
        if (timeoutId) clearTimeout(timeoutId);
        worker.removeAllListeners();
        await worker.terminate();
        resolve(result);
      };

      worker.on("message", async (message) => {
        if (settled) return;

        if (message?.type === "console") {
          logs.push({ level: message.level, message: message.message });
          return;
        }

        if (message?.type === "rpc") {
          try {
            const result = await this.handleRpc(message.method, message.params);
            if (settled) return;
            worker.postMessage({ type: "rpcResult", id: message.id, result });
          } catch (err) {
            const serialized = serializeError(err);
            if (settled) return;
            worker.postMessage({ type: "rpcError", id: message.id, error: serialized });
          }
          return;
        }

        if (message?.type === "result") {
          await settle({ logs });
          return;
        }

        if (message?.type === "error") {
          await settle({ logs, error: message.error });
        }
      });

      worker.on("error", async (err) => {
        await settle({ logs, error: serializeError(err) });
      });

      if (Number.isFinite(timeoutMs) && timeoutMs > 0) {
        timeoutId = setTimeout(() => {
          void settle({
            logs,
            error: {
              name: "ScriptTimeoutError",
              message: `Script timed out after ${timeoutMs}ms`,
            },
          });
        }, timeoutMs);
      }
    });

    worker.postMessage({ type: "run", code });

    return completion;
  }

  async handleRpc(method, params) {
    switch (method) {
      case "ui.alert": {
        throw new Error("ui.alert is not available in the Node ScriptRuntime");
      }
      case "ui.confirm": {
        throw new Error("ui.confirm is not available in the Node ScriptRuntime");
      }
      case "ui.prompt": {
        throw new Error("ui.prompt is not available in the Node ScriptRuntime");
      }
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
