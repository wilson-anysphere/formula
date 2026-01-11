import { PyodideRuntime } from "@formula/python-runtime";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";
import { applyMacroCellUpdates } from "../../macros/applyUpdates";

const PYODIDE_INDEX_URL = globalThis.__pyodideIndexURL || "/pyodide/v0.25.1/full/";
const DEFAULT_NATIVE_PERMISSIONS = { filesystem: "none", network: "none" };
const DEFAULT_TIMEOUT_MS = 5_000;
const DEFAULT_MAX_MEMORY_BYTES = 256 * 1024 * 1024;

/**
 * @param {any[] | undefined} raw
 */
function normalizeUpdates(raw) {
  if (!Array.isArray(raw) || raw.length === 0) return [];
  const out = [];
  for (const u of raw) {
    if (!u || typeof u !== "object") continue;
    const sheetId = String(u.sheet_id ?? "").trim();
    const row = Number(u.row);
    const col = Number(u.col);
    if (!sheetId) continue;
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    out.push({
      sheetId,
      row,
      col,
      value: u.value ?? null,
      formula: typeof u.formula === "string" ? u.formula : null,
      displayValue: String(u.display_value ?? ""),
    });
  }
  return out;
}

function getTauriInvoke(explicit) {
  if (typeof explicit === "function") return explicit;
  return globalThis.__TAURI__?.core?.invoke;
}

/**
 * @param {{
 *   doc: import("../../document/documentController.js").DocumentController,
 *   container: HTMLElement,
 *   workbookId?: string,
 *   invoke?: (cmd: string, args?: any) => Promise<any>,
 *   drainBackendSync?: () => Promise<void>,
 *   getActiveSheetId?: () => string,
 *   getSelection?: () => { sheet_id: string, start_row: number, start_col: number, end_row: number, end_col: number },
 *   setSelection?: (selection: { sheet_id: string, start_row: number, start_col: number, end_row: number, end_col: number }) => void,
 * }} params
 * @returns {{ dispose: () => void }}
 */
export function mountPythonPanel({
  doc,
  container,
  workbookId,
  invoke,
  drainBackendSync,
  getActiveSheetId,
  getSelection,
  setSelection,
}) {
  const isolation = {
    crossOriginIsolated: globalThis.crossOriginIsolated === true,
    sharedArrayBuffer: typeof SharedArrayBuffer !== "undefined",
  };

  let initPromise = null;
  /** @type {PyodideRuntime | null} */
  let pyodideRuntime = null;
  /** @type {any | null} */
  let pyodideBridge = null;

  const tauriInvoke = getTauriInvoke(invoke);
  const nativeAvailable = typeof tauriInvoke === "function";

  container.innerHTML = "";

  const toolbar = document.createElement("div");
  toolbar.style.display = "flex";
  toolbar.style.gap = "8px";
  toolbar.style.padding = "8px";
  toolbar.style.borderBottom = "1px solid var(--panel-border)";

  const runtimeSelect = document.createElement("select");
  runtimeSelect.dataset.testid = "python-panel-runtime";
  runtimeSelect.style.maxWidth = "240px";
  runtimeSelect.title = "Python runtime";
  const pyodideOption = document.createElement("option");
  pyodideOption.value = "pyodide";
  pyodideOption.textContent = "Pyodide (Web)";
  runtimeSelect.appendChild(pyodideOption);
  if (nativeAvailable) {
    const nativeOption = document.createElement("option");
    nativeOption.value = "native";
    nativeOption.textContent = "Native Python (Desktop)";
    runtimeSelect.appendChild(nativeOption);
  }
  runtimeSelect.value = nativeAvailable ? "native" : "pyodide";
  toolbar.appendChild(runtimeSelect);

  const runButton = document.createElement("button");
  runButton.type = "button";
  runButton.textContent = "Run";
  runButton.dataset.testid = "python-panel-run";
  toolbar.appendChild(runButton);

  const isolationLabel = document.createElement("div");
  isolationLabel.dataset.testid = "python-panel-isolation";
  isolationLabel.style.marginLeft = "auto";
  isolationLabel.style.fontSize = "12px";
  isolationLabel.style.color = "var(--text-secondary)";
  toolbar.appendChild(isolationLabel);

  const degradedBanner = document.createElement("div");
  degradedBanner.style.padding = "8px";
  degradedBanner.style.borderBottom = "1px solid var(--panel-border)";
  degradedBanner.style.fontSize = "12px";
  degradedBanner.style.color = "var(--text-secondary)";
  degradedBanner.style.background = "var(--bg-tertiary)";
  degradedBanner.dataset.testid = "python-panel-degraded-banner";
  degradedBanner.textContent =
    "SharedArrayBuffer unavailable; running Pyodide on main thread (UI may freeze during execution).";
  degradedBanner.style.display = "none";

  const editorHost = document.createElement("div");
  editorHost.style.flex = "1";
  editorHost.style.minHeight = "0";

  const consoleHost = document.createElement("pre");
  consoleHost.style.height = "140px";
  consoleHost.style.margin = "0";
  consoleHost.style.padding = "8px";
  consoleHost.style.overflow = "auto";
  consoleHost.style.borderTop = "1px solid var(--panel-border)";
  consoleHost.dataset.testid = "python-panel-output";
  consoleHost.textContent = "Output…";

  const root = document.createElement("div");
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.height = "100%";

  root.appendChild(toolbar);
  root.appendChild(degradedBanner);
  root.appendChild(editorHost);
  root.appendChild(consoleHost);
  container.appendChild(root);

  const editor = document.createElement("textarea");
  editor.value = defaultScript();
  editor.dataset.testid = "python-panel-code";
  editor.spellcheck = false;
  editor.style.width = "100%";
  editor.style.height = "100%";
  editor.style.resize = "none";
  editor.style.border = "none";
  editor.style.outline = "none";
  editor.style.padding = "8px";
  editor.style.boxSizing = "border-box";
  editor.style.fontFamily =
    'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace';
  editor.style.fontSize = "13px";
  editorHost.appendChild(editor);

  const setOutput = (text) => {
    consoleHost.textContent = text;
  };

  const effectivePyodideBackendMode = () => {
    if (pyodideRuntime && typeof pyodideRuntime.getBackendMode === "function") {
      return pyodideRuntime.getBackendMode();
    }
    const canUseWorker =
      typeof Worker !== "undefined" && isolation.sharedArrayBuffer && isolation.crossOriginIsolated;
    return canUseWorker ? "worker" : "mainThread";
  };

  const updateRuntimeStatus = () => {
    if (runtimeSelect.value === "native") {
      isolationLabel.textContent = nativeAvailable ? "Using system Python via Tauri" : "Native runtime unavailable";
      degradedBanner.style.display = "none";
      return;
    }

    const backendMode = effectivePyodideBackendMode();
    isolationLabel.textContent =
      backendMode === "worker"
        ? "Pyodide worker (SharedArrayBuffer enabled)"
        : "Pyodide main thread (SharedArrayBuffer unavailable)";
    degradedBanner.style.display = backendMode === "mainThread" ? "block" : "none";
  };

  updateRuntimeStatus();

  runtimeSelect.addEventListener("change", () => {
    updateRuntimeStatus();
    setOutput("");
  });

  class PanelBridge extends DocumentControllerBridge {
    constructor(doc, options) {
      super(doc, options);
      this._getActiveSheetId = options.getActiveSheetId;
      this._getSelection = options.getSelection;
      this._setSelection = options.setSelection;
    }

    get_active_sheet_id() {
      const sheetId = this._getActiveSheetId?.();
      if (sheetId) {
        this.activeSheetId = sheetId;
        this.sheetIds.add(sheetId);
        if (this.selection) this.selection.sheet_id = sheetId;
      }
      return this.activeSheetId;
    }

    get_selection() {
      const selection = this._getSelection?.();
      if (selection && selection.sheet_id) {
        this.activeSheetId = selection.sheet_id;
        this.sheetIds.add(selection.sheet_id);
        this.selection = { ...selection };
      }
      return { ...this.selection };
    }

    set_selection({ selection }) {
      if (selection && selection.sheet_id) {
        try {
          this._setSelection?.(selection);
        } catch {
          // ignore
        }
      }
      return super.set_selection({ selection });
    }
  }
  async function ensureInitialized() {
    if (pyodideRuntime) {
      if (initPromise) return await initPromise;
      initPromise = pyodideRuntime.initialize().catch((err) => {
        initPromise = null;
        throw err;
      });
      return await initPromise;
    }
    if (initPromise) return await initPromise;
    pyodideBridge = new PanelBridge(doc, {
      activeSheetId: getActiveSheetId?.() ?? "Sheet1",
      getActiveSheetId,
      getSelection,
      setSelection,
    });
    pyodideRuntime = new PyodideRuntime({
      api: pyodideBridge,
      indexURL: PYODIDE_INDEX_URL,
      rpcTimeoutMs: 5_000,
    });
    updateRuntimeStatus();
    if (initPromise) return await initPromise;
    initPromise = pyodideRuntime.initialize().catch((err) => {
      initPromise = null;
      throw err;
    });
    return await initPromise;
  }

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    setOutput("");

    try {
      if (runtimeSelect.value === "native") {
        if (!nativeAvailable) {
          throw new Error("Native Python runtime is not available (Tauri invoke missing)");
        }

        // Allow microtask-batched backend edits to enqueue, then flush so the backend workbook
        // state matches the grid before running the script (same pattern as macros).
        await new Promise((resolve) => queueMicrotask(resolve));
        await drainBackendSync?.();

        const ctx = {
          active_sheet_id: getActiveSheetId?.() ?? undefined,
          selection: getSelection?.() ?? undefined,
        };

        const result = await tauriInvoke("run_python_script", {
          workbook_id: workbookId ?? null,
          code: editor.value,
          permissions: DEFAULT_NATIVE_PERMISSIONS,
          timeout_ms: DEFAULT_TIMEOUT_MS,
          max_memory_bytes: DEFAULT_MAX_MEMORY_BYTES,
          context: ctx,
        });

        const updates = normalizeUpdates(result?.updates);
        if (updates.length > 0) {
          doc.beginBatch({ label: "Run Python" });
          let committed = false;
          try {
            applyMacroCellUpdates(doc, updates);
            committed = true;
          } finally {
            if (committed) doc.endBatch();
            else doc.cancelBatch();
          }
        }

        const stdout = typeof result?.stdout === "string" ? result.stdout : "";
        const stderr = typeof result?.stderr === "string" ? result.stderr : "";
        const errMessage = result?.error?.message ? String(result.error.message) : "";
        const errStack = result?.error?.stack ? String(result.error.stack) : "";
        const header = errMessage ? `${errMessage}${errStack ? `\n\n${errStack}` : ""}` : "";
        const output = [
          header ? `--- error ---\n${header}` : null,
          stdout ? `--- stdout ---\n${stdout}` : null,
          stderr ? `--- stderr ---\n${stderr}` : null,
        ]
          .filter(Boolean)
          .join("\n\n");

        setOutput(output || "(no output)");
        return;
      }

      await ensureInitialized();
      pyodideBridge.activeSheetId = getActiveSheetId?.() ?? pyodideBridge.activeSheetId;
      pyodideBridge.sheetIds.add(pyodideBridge.activeSheetId);
      if (pyodideBridge.selection) pyodideBridge.selection.sheet_id = pyodideBridge.activeSheetId;
      const result = await pyodideRuntime.execute(editor.value);

      const stdout = typeof result.stdout === "string" ? result.stdout : "";
      const stderr = typeof result.stderr === "string" ? result.stderr : "";
      const output = [stdout ? `--- stdout ---\n${stdout}` : null, stderr ? `--- stderr ---\n${stderr}` : null]
        .filter(Boolean)
        .join("\n\n");

      setOutput(output || "(no output)");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      const stdout = typeof err?.stdout === "string" ? err.stdout : "";
      const stderr = typeof err?.stderr === "string" ? err.stderr : "";
      const details = [stdout ? `--- stdout ---\n${stdout}` : null, stderr ? `--- stderr ---\n${stderr}` : null]
        .filter(Boolean)
        .join("\n\n");
      setOutput(details ? `${message}\n\n${details}` : message);
    } finally {
      runButton.disabled = false;
    }
  });

  return {
    dispose: () => {
      pyodideRuntime?.destroy();
      container.innerHTML = "";
    },
  };
}

function defaultScript() {
  return `import formula

sheet = formula.active_sheet
sheet["A1"] = 1
sheet["A2"] = "=A1*2"
`;
}
