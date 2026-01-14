import { PyodideRuntime } from "@formula/python-runtime";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";
import { normalizeFormulaTextOpt } from "@formula/engine";
import { getTauriInvokeOrNull } from "../../tauri/invoke.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../../collab/permissionGuards.js";
import { ensurePyodideIndexURL, getCachedPyodideIndexURL } from "../../pyodide/pyodideIndexURL.js";
const DEFAULT_NATIVE_PERMISSIONS = { filesystem: "none", network: "none" };
const DEFAULT_TIMEOUT_MS = 5_000;
const DEFAULT_MAX_MEMORY_BYTES = 256 * 1024 * 1024;

function valuesEqual(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a !== "object" || typeof b !== "object") return false;
  try {
    return JSON.stringify(a) === JSON.stringify(b);
  } catch {
    return false;
  }
}

function inputEquals(before, after) {
  return valuesEqual(before.value ?? null, after.value ?? null) && (before.formula ?? null) === (after.formula ?? null);
}

function normalizeFormulaText(formula) {
  if (typeof formula !== "string") return null;
  return normalizeFormulaTextOpt(formula);
}

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
  return getTauriInvokeOrNull();
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
 *   isEditing?: () => boolean,
 *   isReadOnly?: () => boolean,
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
  isEditing,
  isReadOnly,
}) {
  const isolation = {
    crossOriginIsolated: globalThis.crossOriginIsolated === true,
    sharedArrayBuffer: typeof SharedArrayBuffer !== "undefined",
  };

  const abort = typeof AbortController !== "undefined" ? new AbortController() : null;
  const eventSignal = abort?.signal;
  let isRunning = false;
  let lastEditing = (globalThis.__formulaSpreadsheetIsEditing ?? false) === true;
  let lastReadOnly = false;
  let blockedOutputReason = null;

  let initPromise = null;
  let pyodideIndexURL = null;
  let pyodideIndexURLPromise = null;
  /** @type {PyodideRuntime | null} */
  let pyodideRuntime = null;
  /** @type {any | null} */
  let pyodideBridge = null;

  const tauriInvoke = getTauriInvoke(invoke);
  const nativeAvailable = typeof tauriInvoke === "function";

  container.innerHTML = "";

  const toolbar = document.createElement("div");
  toolbar.className = "python-panel-mount__toolbar";

  const runtimeSelect = document.createElement("select");
  runtimeSelect.dataset.testid = "python-panel-runtime";
  runtimeSelect.className = "python-panel-mount__runtime-select";
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
  isolationLabel.className = "python-panel-mount__isolation-label";
  toolbar.appendChild(isolationLabel);

  const degradedBanner = document.createElement("div");
  degradedBanner.className = "python-panel-mount__degraded-banner";
  degradedBanner.dataset.testid = "python-panel-degraded-banner";
  degradedBanner.textContent =
    "SharedArrayBuffer unavailable; running Pyodide on main thread (UI may freeze during execution).";
  degradedBanner.hidden = true;

  const editorHost = document.createElement("div");
  editorHost.className = "python-panel-mount__editor-host";

  const consoleHost = document.createElement("pre");
  consoleHost.className = "python-panel-mount__console";
  consoleHost.dataset.testid = "python-panel-output";
  consoleHost.textContent = "Output…";

  const root = document.createElement("div");
  root.className = "python-panel-mount";

  root.appendChild(toolbar);
  root.appendChild(degradedBanner);
  root.appendChild(editorHost);
  root.appendChild(consoleHost);
  container.appendChild(root);

  const editor = document.createElement("textarea");
  editor.value = defaultScript();
  editor.dataset.testid = "python-panel-code";
  editor.spellcheck = false;
  editor.className = "python-panel-mount__editor";
  editorHost.appendChild(editor);

  const setOutput = (text) => {
    consoleHost.textContent = text;
  };

  const getEditingState = () => {
    if (typeof isEditing === "function") {
      try {
        return Boolean(isEditing());
      } catch {
        return false;
      }
    }
    return lastEditing;
  };

  const getReadOnlyState = () => {
    if (typeof isReadOnly === "function") {
      try {
        return Boolean(isReadOnly());
      } catch {
        return false;
      }
    }
    return lastReadOnly;
  };

  const getBlockedReason = () => {
    if (getReadOnlyState()) return READ_ONLY_SHEET_MUTATION_MESSAGE;
    if (getEditingState()) return "Finish editing to run Python.";
    return null;
  };

  const syncRunButtonDisabledState = () => {
    const reason = getBlockedReason();
    const blocked = Boolean(reason);
    runButton.disabled = isRunning || blocked;
    if (blocked) {
      runButton.title = reason;
      runButton.dataset.blockedReason = reason;
      const current = String(consoleHost.textContent ?? "").trim();
      if (current === "" || current === "Output…" || (blockedOutputReason && current === blockedOutputReason.trim())) {
        setOutput(reason);
        blockedOutputReason = reason;
      }
      return;
    }
    runButton.removeAttribute("title");
    delete runButton.dataset.blockedReason;
    const current = String(consoleHost.textContent ?? "").trim();
    if (blockedOutputReason && current === blockedOutputReason.trim()) {
      setOutput("Output…");
    }
    blockedOutputReason = null;
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
      degradedBanner.hidden = true;
      return;
    }

    const backendMode = effectivePyodideBackendMode();
    isolationLabel.textContent =
      backendMode === "worker"
        ? "SharedArrayBuffer enabled"
        : "SharedArrayBuffer unavailable — running Pyodide on main thread (may freeze UI)";
    degradedBanner.hidden = backendMode !== "mainThread";
  };

  updateRuntimeStatus();
  syncRunButtonDisabledState();

  const formatBytes = (bytes) => {
    if (!Number.isFinite(bytes) || bytes == null) return "";
    const units = ["B", "KiB", "MiB", "GiB"];
    let value = bytes;
    let unit = units[0];
    for (let idx = 0; idx < units.length - 1; idx += 1) {
      if (value < 1024) break;
      value /= 1024;
      unit = units[idx + 1];
    }
    return `${value.toFixed(value >= 10 || unit === "B" ? 0 : 1)} ${unit}`;
  };

  const renderPyodideProgress = (progress) => {
    if (!progress || typeof progress !== "object") return null;
    const kind = progress.kind;
    if (kind === "downloadStart" && progress.message) return String(progress.message);
    if (kind === "ready" && progress.message) return String(progress.message);
    if (kind === "downloadProgress") {
      const name = progress.fileName ? String(progress.fileName) : "asset";
      const total = formatBytes(progress.bytesTotal);
      const done = formatBytes(progress.bytesDownloaded);
      const frac = total ? ` (${done} / ${total})` : done ? ` (${done})` : "";
      return `Downloading ${name}${frac}…`;
    }
    return null;
  };

  async function resolvePyodideIndexURL({ downloadIfMissing }) {
    if (typeof pyodideIndexURL === "string" && pyodideIndexURL.length > 0) return pyodideIndexURL;
    if (pyodideIndexURLPromise) return await pyodideIndexURLPromise;

    pyodideIndexURLPromise = (async () => {
      const resolved = downloadIfMissing
        ? await ensurePyodideIndexURL({
            onProgress: (progress) => {
              const msg = renderPyodideProgress(progress);
              if (msg) setOutput(`${msg}\n`);
            },
          })
        : await getCachedPyodideIndexURL();
      if (typeof resolved === "string" && resolved.length > 0) {
        pyodideIndexURL = resolved;
      }
      return resolved;
    })().finally(() => {
      pyodideIndexURLPromise = null;
    });

    return await pyodideIndexURLPromise;
  }

  runtimeSelect.addEventListener("change", () => {
    updateRuntimeStatus();
    setOutput("");
    if (runtimeSelect.value === "pyodide") {
      void resolvePyodideIndexURL({ downloadIfMissing: true }).catch((err) => {
        const message = err instanceof Error ? err.message : String(err);
        if (runtimeSelect.value === "pyodide") {
          setOutput(`Failed to download Pyodide assets.\n\n${message}`);
        }
      });
    }
  });

  if (typeof window !== "undefined") {
    window.addEventListener(
      "formula:spreadsheet-editing-changed",
      (evt) => {
        lastEditing = Boolean(evt?.detail?.isEditing);
        syncRunButtonDisabledState();
      },
      eventSignal ? { signal: eventSignal } : undefined,
    );
    window.addEventListener(
      "formula:read-only-changed",
      (evt) => {
        lastReadOnly = Boolean(evt?.detail?.readOnly);
        syncRunButtonDisabledState();
      },
      eventSignal ? { signal: eventSignal } : undefined,
    );
  }

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
    if (!pyodideRuntime) {
      pyodideBridge = new PanelBridge(doc, {
        activeSheetId: getActiveSheetId?.() ?? "Sheet1",
        getActiveSheetId,
        getSelection,
        setSelection,
      });
      pyodideRuntime = new PyodideRuntime({
        api: pyodideBridge,
        rpcTimeoutMs: 5_000,
      });
      updateRuntimeStatus();
    }

    if (initPromise) return await initPromise;

    // `PyodideRuntime.initialize()` is idempotent and will return quickly when
    // already initialized. We only memoize the in-flight promise so callers can
    // retry after the runtime resets itself (timeouts/memory errors).
    initPromise = (async () => {
      const indexURL = await resolvePyodideIndexURL({ downloadIfMissing: true });
      await pyodideRuntime.initialize({ indexURL });
    })().finally(() => {
      initPromise = null;
      updateRuntimeStatus();
    });

    return await initPromise;
  }

  runButton.addEventListener("click", () => {
    void (async () => {
      const blockedReason = getBlockedReason();
      if (blockedReason) {
        setOutput(blockedReason);
        try {
          syncRunButtonDisabledState();
        } catch {
          // ignore UI sync failures
        }
        return;
      }

      isRunning = true;
      syncRunButtonDisabledState();
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
            const deltas = [];
            for (const update of updates) {
              // Avoid resurrecting deleted sheets: only apply updates when the sheet still exists.
              const sheetId = String(update.sheetId ?? "").trim();
              if (!sheetId) continue;
              const meta = typeof doc.getSheetMeta === "function" ? doc.getSheetMeta(sheetId) : null;
              if (!meta) {
                const ids = typeof doc.getSheetIds === "function" ? doc.getSheetIds() : [];
                if (Array.isArray(ids) && ids.length > 0) continue;
              }

              const before =
                typeof doc.peekCell === "function"
                  ? doc.peekCell(sheetId, { row: update.row, col: update.col })
                  : doc.getCell(sheetId, { row: update.row, col: update.col });
              const formula = normalizeFormulaText(update.formula);
              const value = formula ? null : (update.value ?? null);
              const after = { value, formula, styleId: before.styleId ?? 0 };
              if (inputEquals(before, after)) continue;
              deltas.push({ sheetId, row: update.row, col: update.col, before, after });
            }
            if (deltas.length > 0) {
              // Native Python scripts mutate the backend workbook directly and return
              // updates for the UI. Apply them without creating a new undo step, and
              // tag them so the workbook sync bridge doesn't echo them back.
              doc.applyExternalDeltas(deltas, { source: "python" });
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
        isRunning = false;
        try {
          syncRunButtonDisabledState();
        } catch {
          // ignore UI sync failures
        }
      }
    })().catch((err) => {
      // Terminal catch to avoid unhandled rejections from click handlers.
      try {
        const message = err instanceof Error ? err.message : String(err);
        setOutput(message);
      } catch {
        // ignore output failures
      }
      isRunning = false;
      try {
        syncRunButtonDisabledState();
      } catch {
        // ignore UI sync failures
      }
    });
  });

  return {
    dispose: () => {
      abort?.abort();
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
