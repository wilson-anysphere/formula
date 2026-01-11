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
const DEFAULT_TIMEOUT_MS = 5_000;

function createRunToken() {
  const cryptoObj = globalThis.crypto;
  if (cryptoObj?.randomUUID) return cryptoObj.randomUUID();

  const bytes = new Uint8Array(16);
  if (cryptoObj?.getRandomValues) {
    cryptoObj.getRandomValues(bytes);
  } else {
    for (let i = 0; i < bytes.length; i += 1) {
      bytes[i] = Math.floor(Math.random() * 256);
    }
  }

  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

export class ScriptRuntime {
  /**
   * @param {Workbook} workbook
   */
  constructor(workbook) {
    this.workbook = workbook;
  }

  /**
   * @param {string} code
   * @param {{ permissions?: ScriptPermissions, timeoutMs?: number }=} options
   * @returns {Promise<ScriptRunResult>}
   *
   * If execution exceeds `timeoutMs`, the worker is terminated and the promise
   * resolves with a `ScriptRunResult` containing an error.
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
    const token = createRunToken();
    const timeoutMs =
      Number.isFinite(options?.timeoutMs) && options.timeoutMs > 0 ? options.timeoutMs : DEFAULT_TIMEOUT_MS;

    /** @type {MessagePort | null} */
    let controlPort = null;
    /** @type {MessagePort | null} */
    let workerPort = null;
    if (typeof MessageChannel !== "undefined") {
      const channel = new MessageChannel();
      controlPort = channel.port1;
      workerPort = channel.port2;
    }

    const completion = new Promise((resolve) => {
      let timeoutId;
      let settled = false;

      const cleanup = () => {
        if (controlPort) {
          controlPort.onmessage = null;
          controlPort.close();
        }
        worker.onmessage = null;
        worker.onerror = null;
        worker.terminate();
      };

      const settle = (result) => {
        if (settled) return;
        settled = true;
        if (timeoutId) clearTimeout(timeoutId);
        cleanup();
        resolve(result);
      };

      const onMessage = async (event) => {
        if (settled) return;
        const message = event.data;
        if (!message || message.token !== token) return;

        if (message.type === "console") {
          logs.push({ level: message.level, message: message.message });
          return;
        }

        if (message.type === "rpc") {
          try {
            const result = await this.handleRpc(message.method, message.params);
            if (settled) return;
            (controlPort ?? worker).postMessage({ type: "rpcResult", token, id: message.id, result });
          } catch (err) {
            if (settled) return;
            (controlPort ?? worker).postMessage({ type: "rpcError", token, id: message.id, error: serializeError(err) });
          }
          return;
        }

        if (message.type === "result") {
          settle({ logs });
          return;
        }

        if (message.type === "error") {
          settle({ logs, error: message.error });
        }
      };

      if (controlPort) {
        controlPort.onmessage = onMessage;
      } else {
        worker.onmessage = onMessage;
      }

      worker.onerror = (event) => {
        settle({ logs, error: serializeError(event?.error ?? event?.message ?? "Worker error") });
      };

      timeoutId = setTimeout(() => {
        settle({
          logs,
          error: {
            name: "ScriptTimeoutError",
            message: `Script timed out after ${timeoutMs}ms`,
          },
        });
      }, timeoutMs);
    });

    worker.postMessage({
      type: "run",
      token,
      code,
      activeSheetName,
      selection,
      permissions: options?.permissions,
      ...(workerPort ? { controlPort: workerPort } : null),
    }, workerPort ? [workerPort] : undefined);

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
