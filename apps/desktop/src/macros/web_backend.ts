import { ScriptRuntime, type PermissionSnapshot as ScriptPermissionSnapshot } from "@formula/scripting/web";
import { PyodideRuntime } from "@formula/python-runtime/pyodide";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

import { DocumentControllerWorkbookAdapter } from "../scripting/documentControllerWorkbookAdapter.js";

import type { DocumentController } from "../document/documentController.js";
import type {
  MacroBackend,
  MacroInfo,
  MacroLanguage,
  MacroPermission,
  MacroRunRequest,
  MacroRunResult,
  MacroSecurityStatus,
  MacroTrustDecision,
} from "./types";

type PythonPermissions = {
  filesystem?: "none" | "read" | "readwrite";
  network?: "none" | "allowlist" | "full";
  networkAllowlist?: string[];
};

type StoredMacro = MacroInfo & { code: string };

export interface WebMacroBackendOptions {
  getDocumentController: () => DocumentController;
  getActiveSheetId?: () => string;
  storage?: Storage;
  /**
   * Base URL to load Pyodide assets from (must end with a trailing slash).
   *
   * For the Vite demo we proxy `/pyodide/**` to jsdelivr so we can run under
   * `crossOriginIsolated` without loading cross-origin scripts directly.
   */
  pyodideIndexURL?: string;
}

const STORAGE_PREFIX = "formula:macros:";

const BUILTIN_MACROS: StoredMacro[] = [
  {
    id: "demo-typescript-write-cell",
    name: "TypeScript: Write E1",
    language: "typescript",
    module: "builtins",
    code: `
await ctx.activeSheet.getRange("E1").setValue("hello from ts");
ctx.ui.log("TypeScript macro wrote E1");
`,
  },
  {
    id: "demo-python-write-cell",
    name: "Python: Write E2",
    language: "python",
    module: "builtins",
    code: `
import formula

sheet = formula.active_sheet
sheet["E2"] = "hello from python"
print("Python macro wrote E2")
`,
  },
];

function storageKey(workbookId: string): string {
  return `${STORAGE_PREFIX}${workbookId}`;
}

function isMacroLanguage(value: unknown): value is MacroLanguage {
  return value === "vba" || value === "typescript" || value === "python";
}

function normalizeStoredMacro(raw: unknown): StoredMacro | null {
  if (!raw || typeof raw !== "object") return null;
  const obj = raw as Record<string, unknown>;
  const id = typeof obj.id === "string" ? obj.id : null;
  const name = typeof obj.name === "string" ? obj.name : null;
  const language = obj.language;
  const code = typeof obj.code === "string" ? obj.code : null;
  const module = typeof obj.module === "string" ? obj.module : undefined;
  if (!id || !name || !code || !isMacroLanguage(language)) return null;
  return { id, name, language, code, module };
}

function formatConsoleEntry(entry: { level: string; message: string }): string {
  if (entry.level === "log") return entry.message;
  return `${entry.level}: ${entry.message}`;
}

function splitLines(text: unknown): string[] {
  if (typeof text !== "string" || text.length === 0) return [];
  const lines = text.split(/\r?\n/);
  // Drop trailing newline.
  if (lines.length > 0 && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

function prefixLines(prefix: string, text: unknown): string[] {
  return splitLines(text).map((line) => `${prefix}${line}`);
}

function defaultScriptPermissions(): ScriptPermissionSnapshot {
  return { network: { mode: "none", allowlist: [] } };
}

function defaultPythonPermissions(): PythonPermissions {
  return { filesystem: "none", network: "none" };
}

function effectiveNetworkAllowlist(): string[] {
  const hostname = globalThis.location?.hostname;
  return hostname ? [hostname] : ["localhost"];
}

function scriptPermissionsFromMacroPermissions(perms: MacroPermission[] | undefined): ScriptPermissionSnapshot {
  const requested = new Set(perms ?? []);
  if (!requested.has("network")) return defaultScriptPermissions();
  return { network: { mode: "allowlist", allowlist: effectiveNetworkAllowlist() } };
}

function pythonPermissionsFromMacroPermissions(perms: MacroPermission[] | undefined): PythonPermissions {
  const requested = new Set(perms ?? []);

  let filesystem: PythonPermissions["filesystem"] = "none";
  if (requested.has("filesystem_write")) filesystem = "readwrite";
  else if (requested.has("filesystem_read")) filesystem = "read";

  if (!requested.has("network")) return { filesystem, network: "none" };
  return { filesystem, network: "allowlist", networkAllowlist: effectiveNetworkAllowlist() };
}

function macroError(err: unknown): { message: string; stack?: string } {
  if (err instanceof Error) {
    return { message: err.message, stack: err.stack };
  }
  return { message: String(err) };
}

function canUsePyodideWorkerBackend(): boolean {
  return (
    typeof Worker !== "undefined" &&
    typeof SharedArrayBuffer !== "undefined" &&
    globalThis.crossOriginIsolated === true
  );
}

export class WebMacroBackend implements MacroBackend {
  private readonly storage: Storage | null;
  private readonly macrosCache = new Map<string, StoredMacro[]>();

  private pyodide: PyodideRuntime | null = null;
  private pyodideInit: Promise<void> | null = null;
  private warnedPyodideMainThread = false;

  constructor(private readonly options: WebMacroBackendOptions) {
    this.storage =
      options.storage ??
      (() => {
        try {
          return globalThis.localStorage ?? null;
        } catch {
          return null;
        }
      })();
  }

  async listMacros(workbookId: string): Promise<MacroInfo[]> {
    // Warm up Pyodide in the background so the first Python run doesn't pay the
    // full initialization cost.
    void this.prewarmPython().catch(() => {
      // ignore – the UI will surface errors when the user runs a Python macro.
    });

    return this.getMacros(workbookId, { refresh: true }).map(({ code: _code, ...info }) => info);
  }

  async getMacroSecurityStatus(_workbookId: string): Promise<MacroSecurityStatus> {
    // Web builds do not integrate with the desktop Trust Center. Treat workbook macros
    // as absent so the macros UI can still function for TypeScript/Python demos.
    return { hasMacros: false, trust: "blocked" };
  }

  async setMacroTrust(workbookId: string, _decision: MacroTrustDecision): Promise<MacroSecurityStatus> {
    return await this.getMacroSecurityStatus(workbookId);
  }

  async runMacro(request: MacroRunRequest): Promise<MacroRunResult> {
    const macro = this.getMacros(request.workbookId, { refresh: true }).find((m) => m.id === request.macroId);
    if (!macro) {
      return { ok: false, output: [], error: { message: `Unknown macro id: ${request.macroId}` } };
    }

    switch (macro.language) {
      case "typescript":
        return await this.runTypeScriptMacro(macro, request);
      case "python":
        return await this.runPythonMacro(macro, request);
      case "vba":
        return { ok: false, output: [], error: { message: "VBA macros are not supported in the web demo." } };
      default:
        return { ok: false, output: [], error: { message: `Unsupported macro language: ${String(macro.language)}` } };
    }
  }

  private activeSheetId(): string {
    return this.options.getActiveSheetId?.() ?? "Sheet1";
  }

  private getMacros(workbookId: string): StoredMacro[];
  private getMacros(workbookId: string, options: { refresh?: boolean }): StoredMacro[];
  private getMacros(workbookId: string, options: { refresh?: boolean } = {}): StoredMacro[] {
    const refresh = options.refresh === true;
    if (!refresh) {
      const cached = this.macrosCache.get(workbookId);
      if (cached) return cached;
    }

    const stored = this.readMacrosFromStorage(workbookId);
    const byId = new Map<string, StoredMacro>();
    for (const macro of stored) byId.set(macro.id, macro);
    for (const builtin of BUILTIN_MACROS) {
      if (!byId.has(builtin.id)) byId.set(builtin.id, builtin);
    }

    const macros = Array.from(byId.values());
    macros.sort((a, b) => a.name.localeCompare(b.name));
    this.macrosCache.set(workbookId, macros);

    // Persist the merged set if we have storage and were missing built-ins.
    if (this.storage) {
      const missingBuiltins = BUILTIN_MACROS.some((builtin) => !stored.some((m) => m.id === builtin.id));
      if (missingBuiltins || stored.length === 0) {
        this.writeMacrosToStorage(workbookId, macros);
      }
    }

    return macros;
  }

  private readMacrosFromStorage(workbookId: string): StoredMacro[] {
    if (!this.storage) return [];
    try {
      const raw = this.storage.getItem(storageKey(workbookId));
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      const macros: StoredMacro[] = [];
      for (const entry of parsed) {
        const normalized = normalizeStoredMacro(entry);
        if (normalized) macros.push(normalized);
      }
      return macros;
    } catch {
      return [];
    }
  }

  private writeMacrosToStorage(workbookId: string, macros: StoredMacro[]): void {
    if (!this.storage) return;
    try {
      this.storage.setItem(storageKey(workbookId), JSON.stringify(macros));
    } catch {
      // Ignore persistence failures (e.g. storage disabled).
    }
  }

  private async runTypeScriptMacro(macro: StoredMacro, request: MacroRunRequest): Promise<MacroRunResult> {
    const doc = this.options.getDocumentController();
    const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: this.activeSheetId() });
    const runtime = new ScriptRuntime(workbook as any);

    try {
      const permissions = scriptPermissionsFromMacroPermissions(request.permissions);
      const result = await runtime.run(macro.code, { permissions, timeoutMs: request.timeoutMs });

      const output = Array.isArray(result.logs) ? result.logs.map(formatConsoleEntry) : [];
      if (result.error) {
        return {
          ok: false,
          output,
          error: { message: String(result.error.message ?? "Script failed"), stack: result.error.stack },
        };
      }

      return { ok: true, output };
    } catch (err) {
      return { ok: false, output: [], error: macroError(err) };
    } finally {
      workbook.dispose();
    }
  }

  private async runPythonMacro(macro: StoredMacro, request: MacroRunRequest): Promise<MacroRunResult> {
    const doc = this.options.getDocumentController();
    const api = new DocumentControllerBridge(doc, { activeSheetId: this.activeSheetId() });

    const output: string[] = [];

    try {
      const runtime = await this.ensurePyodideInitialized(api);
      if (runtime.getBackendMode?.() === "mainThread" && !this.warnedPyodideMainThread) {
        output.push(
          "warning: SharedArrayBuffer unavailable; running Pyodide on the main thread (UI may freeze during execution).",
        );
        this.warnedPyodideMainThread = true;
      }
      const permissions = pythonPermissionsFromMacroPermissions(request.permissions);
      const result = await runtime.execute(macro.code, { timeoutMs: request.timeoutMs, permissions });

      output.push(...prefixLines("stdout: ", result?.stdout));
      output.push(...prefixLines("stderr: ", result?.stderr));

      return { ok: true, output };
    } catch (err) {
      // PyodideRuntime attaches stdout/stderr to thrown errors as best-effort.
      output.push(...prefixLines("stdout: ", (err as any)?.stdout));
      output.push(...prefixLines("stderr: ", (err as any)?.stderr));
      return { ok: false, output, error: macroError(err) };
    }
  }

  private async prewarmPython(): Promise<void> {
    // Prewarming Pyodide on the main thread can freeze the UI. Only warm up
    // eagerly when we can use the Worker backend.
    if (!canUsePyodideWorkerBackend()) return;

    const doc = this.options.getDocumentController();
    const api = new DocumentControllerBridge(doc, { activeSheetId: this.activeSheetId() });
    await this.ensurePyodideInitialized(api);
  }

  private async ensurePyodideInitialized(api: unknown): Promise<PyodideRuntime> {
    const indexURL =
      this.options.pyodideIndexURL ??
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ((globalThis as any).__pyodideIndexURL as string | undefined) ??
      "/pyodide/v0.25.1/full/";

    if (!this.pyodide) {
      this.pyodide = new PyodideRuntime({
        api,
        indexURL,
        mode: "auto",
        permissions: defaultPythonPermissions(),
        timeoutMs: 5_000,
        maxMemoryBytes: 256 * 1024 * 1024,
      });
    }

    // Both backends read `this.api` at call time, so updating it here swaps the
    // bridge between executions (used to keep `formula.active_sheet` aligned).
    (this.pyodide as any).api = api;

    if ((this.pyodide as any).initialized !== true) {
      this.pyodideInit = null;
    }

    if (!this.pyodideInit) {
      this.pyodideInit = this.pyodide.initialize({
        api,
        indexURL,
        permissions: defaultPythonPermissions(),
      });
    }

    try {
      await this.pyodideInit;
    } catch (err) {
      const runtimeErr = err instanceof Error ? err : new Error(String(err));
      const needsIsolation = globalThis.crossOriginIsolated !== true || typeof (globalThis as any).SharedArrayBuffer === "undefined";
      const guidance = needsIsolation
        ? "SharedArrayBuffer is unavailable, so Pyodide is running on the main thread.\n" +
          "If initialization fails, ensure Pyodide assets are reachable (pyodideIndexURL) or enable COOP/COEP headers to use the worker backend."
        : "If initialization fails, ensure Pyodide assets are reachable (pyodideIndexURL).";

      runtimeErr.message = `${runtimeErr.message}\n\n${guidance}`;

      // Reset so a subsequent attempt can retry initialization.
      this.pyodide?.destroy();
      this.pyodide = null;
      this.pyodideInit = null;
      throw runtimeErr;
    }

    return this.pyodide;
  }
}
