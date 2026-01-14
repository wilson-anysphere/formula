import { Worker } from "node:worker_threads";

/**
 * Node ScriptRuntime powered by a hardened sandbox (via @formula/security).
 *
 * This file is only exported from `@formula/scripting/node`.
 */

/**
 * @typedef {{ level: "log" | "info" | "warn" | "error", message: string }} ScriptConsoleEntry
 * @typedef {{ log: (event: any) => void }} AuditSink
 * @typedef {{ type: string, id: string }} ScriptPrincipal
 * @typedef {{
 *   filesystem?: { read?: string[], readwrite?: string[] },
 *   network?: { mode?: "none" | "allowlist" | "full", allowlist?: string[] },
 *   clipboard?: boolean,
 *   notifications?: boolean,
 *   automation?: boolean,
 * }} PermissionSnapshot
 * @typedef {{
 *   logs: ScriptConsoleEntry[],
 *   audit: any[],
 *   error?: { name?: string, message: string, stack?: string, code?: string, principal?: any, request?: any, reason?: string }
 * }} ScriptRunResult
 * @typedef {import("./workbook.js").Workbook} Workbook
 */

const WORKER_URL = new URL("./sandbox-worker.node.js", import.meta.url);

function defaultPermissionSnapshot() {
  return {
    filesystem: { read: [], readwrite: [] },
    network: { mode: "none", allowlist: [] },
    clipboard: false,
    notifications: false,
    automation: false,
  };
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
   * @param {{
   *   timeoutMs?: number,
   *   memoryMb?: number,
   *   signal?: AbortSignal,
   *   principal?: ScriptPrincipal,
   *   permissions?: PermissionSnapshot,
   *   permissionManager?: { getSnapshot: (principal: ScriptPrincipal) => PermissionSnapshot },
   *   auditSink?: AuditSink,
   * }} [options]
   * @returns {Promise<ScriptRunResult>}
   */
  async run(code, options = {}) {
    const activeSheetName = this.workbook.getActiveSheet().name;
    const selection = this.workbook.getSelection();

    /** @type {ScriptConsoleEntry[]} */
    const logs = [];
    /** @type {any[]} */
    const audit = [];

    // Default to a more forgiving timeout since script execution includes worker
    // startup + TypeScript transpilation (which can vary depending on load).
    //
    // Note: CI/vitest runs can be heavily parallelized, so worker startup can occasionally
    // exceed a 20s budget even for small scripts. Prefer a slightly larger default and let
    // callers opt into stricter limits via `options.timeoutMs`.
    const timeoutMs = options.timeoutMs ?? 40_000;
    const memoryMb = options.memoryMb ?? 64;
    const principal = options.principal ?? { type: "script", id: "anonymous" };
    const permissions =
      options.permissions ??
      (options.permissionManager ? options.permissionManager.getSnapshot(principal) : defaultPermissionSnapshot());

    const worker = new Worker(WORKER_URL, {
      type: "module",
      resourceLimits: {
        maxOldGenerationSizeMb: Math.max(16, Math.floor(memoryMb)),
        maxYoungGenerationSizeMb: Math.max(16, Math.min(64, Math.floor(memoryMb / 4))),
      },
    });

    const unsubscribes = [
      this.workbook.events.on("cellChanged", (evt) => {
        worker.postMessage({ type: "event", eventType: "edit", payload: evt });
      }),
      this.workbook.events.on("formulaChanged", (evt) => {
        worker.postMessage({ type: "event", eventType: "edit", payload: evt });
      }),
      this.workbook.events.on("selectionChanged", (evt) => {
        worker.postMessage({ type: "event", eventType: "selectionChange", payload: evt });
      }),
      this.workbook.events.on("formatChanged", (evt) => {
        worker.postMessage({ type: "event", eventType: "formatChange", payload: evt });
      }),
    ];

    const completion = new Promise((resolve) => {
      let settled = false;

      const cleanup = async () => {
        for (const unsub of unsubscribes) {
          try {
            unsub();
          } catch {
            // ignore cleanup failures
          }
        }
        try {
          worker.removeAllListeners();
        } catch {
          // ignore
        }
        try {
          await worker.terminate();
        } catch {
          // ignore
        }
      };

      const settle = async (result) => {
        if (settled) return;
        settled = true;
        try {
          clearTimeout(timeout);
        } catch {
          // ignore
        }
        if (abortListener) {
          try {
            options.signal?.removeEventListener("abort", abortListener);
          } catch {
            // ignore
          }
        }
        try {
          await cleanup();
        } catch {
          // ignore cleanup failures
        }
        resolve(result);
      };

      const onTimeout = () => {
        audit.push({ eventType: "scripting.run.timeout", actor: principal, success: false, metadata: { timeoutMs } });
        try {
          options.auditSink?.log?.({
            eventType: "scripting.run.timeout",
            actor: principal,
            success: false,
            metadata: { timeoutMs },
          });
        } catch {
          // ignore audit failures
        }
        try {
          worker.postMessage({ type: "cancel" });
        } catch {
          // ignore cancellation failures
        }
        void settle({
          logs,
          audit,
          error: { name: "SandboxTimeoutError", message: `Script timed out after ${timeoutMs}ms` },
        }).catch(() => {});
      };

      const timeout = Number.isFinite(timeoutMs) && timeoutMs > 0 ? setTimeout(onTimeout, timeoutMs) : null;

      /** @type {(() => void) | null} */
      let abortListener = null;
      if (options.signal) {
        if (options.signal.aborted) {
          try {
            worker.postMessage({ type: "cancel" });
          } catch {
            // ignore cancellation failures
          }
          void settle({ logs, audit, error: { name: "AbortError", message: "Script execution cancelled" } }).catch(
            () => {},
          );
        } else {
          abortListener = () => {
            audit.push({ eventType: "scripting.run.cancelled", actor: principal, success: false, metadata: {} });
            try {
              options.auditSink?.log?.({
                eventType: "scripting.run.cancelled",
                actor: principal,
                success: false,
                metadata: {},
              });
            } catch {
              // ignore audit failures
            }
            try {
              worker.postMessage({ type: "cancel" });
            } catch {
              // ignore cancellation failures
            }
            void settle({ logs, audit, error: { name: "AbortError", message: "Script execution cancelled" } }).catch(
              () => {},
            );
          };
          options.signal.addEventListener("abort", abortListener, { once: true });
        }
      }

      worker.on("message", (message) => {
        void (async () => {
          if (settled) return;

          if (message?.type === "console") {
            logs.push({ level: message.level, message: message.message });
            return;
          }

          if (message?.type === "output") {
            const level = message.metadata?.method ?? (message.stream === "stderr" ? "error" : "log");
            const text = String(message.text ?? "");
            for (const line of text.split("\n")) {
              const trimmed = line.trimEnd();
              if (!trimmed) continue;
              logs.push({ level, message: trimmed });
            }
            return;
          }

          if (message?.type === "audit") {
            audit.push(message.event);
            try {
              options.auditSink?.log?.(message.event);
            } catch {
              // ignore audit failures
            }
            return;
          }

          if (message?.type === "limit") {
            const event = {
              eventType: "scripting.sandbox.limit",
              actor: principal,
              success: false,
              metadata: message,
            };
            audit.push(event);
            try {
              options.auditSink?.log?.(event);
            } catch {
              // ignore audit failures
            }
            return;
          }

          if (message?.type === "rpc") {
            const started = Date.now();
            try {
              const result = await this.handleRpc(message.method, message.params);
              if (settled) return;
              try {
                worker.postMessage({ type: "rpcResult", id: message.id, result });
              } catch {
                // ignore response post failures (worker may already be shutting down)
              }
              const durationMs = Date.now() - started;
              const entry = {
                eventType: "scripting.rpc",
                actor: principal,
                success: true,
                metadata: { method: message.method, durationMs },
              };
              audit.push(entry);
              try {
                options.auditSink?.log?.(entry);
              } catch {
                // ignore audit failures
              }
            } catch (err) {
              const serialized = serializeError(err);
              if (settled) return;
              try {
                worker.postMessage({ type: "rpcError", id: message.id, error: serialized });
              } catch {
                // ignore response post failures (worker may already be shutting down)
              }
              const durationMs = Date.now() - started;
              const entry = {
                eventType: "scripting.rpc",
                actor: principal,
                success: false,
                metadata: { method: message.method, durationMs, message: serialized.message },
              };
              audit.push(entry);
              try {
                options.auditSink?.log?.(entry);
              } catch {
                // ignore audit failures
              }
            }
            return;
          }

          if (message?.type === "result") {
            await settle({ logs, audit });
            return;
          }

          if (message?.type === "error") {
            await settle({ logs, audit, error: message.error });
          }
        })().catch((err) => {
          void settle({ logs, audit, error: serializeError(err) }).catch(() => {});
        });
      });

      worker.on("error", (err) => {
        void settle({ logs, audit, error: serializeError(err) }).catch(() => {});
      });
    });

    worker.postMessage({
      type: "run",
      code,
      activeSheetName,
      selection,
      principal,
      permissions,
      timeoutMs,
      memoryMb,
    });

    return completion;
  }

  async handleRpc(method, params) {
    switch (method) {
      case "ui.alert":
      case "ui.confirm":
      case "ui.prompt": {
        throw new Error(`${method} is not available in the Node ScriptRuntime`);
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
      case "range.getFormulas": {
        const { sheetName, address } = params;
        return this.workbook.getSheet(sheetName).getRange(address).getFormulas();
      }
      case "range.setFormulas": {
        const { sheetName, address, formulas } = params;
        this.workbook.getSheet(sheetName).getRange(address).setFormulas(formulas);
        return null;
      }
      case "range.getFormats": {
        const { sheetName, address } = params;
        return this.workbook.getSheet(sheetName).getRange(address).getFormats();
      }
      case "range.setFormats": {
        const { sheetName, address, formats } = params;
        this.workbook.getSheet(sheetName).getRange(address).setFormats(formats);
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
      case "workbook.getSheets": {
        return this.workbook.getSheets().map((sheet) => sheet.name);
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
      case "sheet.getUsedRange": {
        const { sheetName } = params;
        return this.workbook.getSheet(sheetName).getUsedRange().address;
      }
      default:
        throw new Error(`Unknown RPC method: ${method}`);
    }
  }
}

function serializeError(err) {
  if (err instanceof Error) {
    return { message: err.message, name: err.name, stack: err.stack, code: err.code };
  }
  if (typeof err === "string") {
    return { message: err };
  }
  return { message: "Unknown error" };
}
