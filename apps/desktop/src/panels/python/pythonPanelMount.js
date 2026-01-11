import { PyodideRuntime } from "@formula/python-runtime";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

const PYODIDE_INDEX_URL = globalThis.__pyodideIndexURL || "/pyodide/v0.25.1/full/";

/**
 * @param {{
 *   doc: import("../../document/documentController.js").DocumentController,
 *   container: HTMLElement,
 *   getActiveSheetId?: () => string,
 *   getSelection?: () => { sheet_id: string, start_row: number, start_col: number, end_row: number, end_col: number },
 *   setSelection?: (selection: { sheet_id: string, start_row: number, start_col: number, end_row: number, end_col: number }) => void,
 * }} params
 * @returns {{ dispose: () => void }}
 */
export function mountPythonPanel({ doc, container, getActiveSheetId, getSelection, setSelection }) {
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

  const bridge = new PanelBridge(doc, {
    activeSheetId: getActiveSheetId?.() ?? "Sheet1",
    getActiveSheetId,
    getSelection,
    setSelection,
  });

  const runtime = new PyodideRuntime({
    api: bridge,
    indexURL: PYODIDE_INDEX_URL,
    rpcTimeoutMs: 5_000,
  });

  const isolation = {
    crossOriginIsolated: globalThis.crossOriginIsolated === true,
    sharedArrayBuffer: typeof SharedArrayBuffer !== "undefined",
  };

  let initPromise = null;

  container.innerHTML = "";

  const toolbar = document.createElement("div");
  toolbar.style.display = "flex";
  toolbar.style.gap = "8px";
  toolbar.style.padding = "8px";
  toolbar.style.borderBottom = "1px solid var(--panel-border)";

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
  isolationLabel.textContent = isolation.sharedArrayBuffer
    ? "SharedArrayBuffer enabled"
    : "SharedArrayBuffer unavailable (crossOriginIsolated required)";
  toolbar.appendChild(isolationLabel);

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

  if (!isolation.sharedArrayBuffer || !isolation.crossOriginIsolated) {
    setOutput(
      "SharedArrayBuffer is required for the Pyodide formula bridge.\n\n" +
        "In browsers/webviews this requires a cross-origin isolated context (COOP/COEP).\n" +
        "Formula's Vite dev server config enables this automatically; other hosts must do the same.",
    );
  }

  async function ensureInitialized() {
    if (initPromise) return await initPromise;
    initPromise = runtime.initialize().catch((err) => {
      initPromise = null;
      throw err;
    });
    return await initPromise;
  }

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    setOutput("");

    try {
      bridge.activeSheetId = getActiveSheetId?.() ?? bridge.activeSheetId;
      bridge.sheetIds.add(bridge.activeSheetId);
      if (bridge.selection) bridge.selection.sheet_id = bridge.activeSheetId;

      await ensureInitialized();
      const result = await runtime.execute(editor.value);

      const stdout = typeof result.stdout === "string" ? result.stdout : "";
      const stderr = typeof result.stderr === "string" ? result.stderr : "";
      const output = [
        stdout ? `--- stdout ---\n${stdout}` : null,
        stderr ? `--- stderr ---\n${stderr}` : null,
      ]
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
      runtime.destroy();
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
