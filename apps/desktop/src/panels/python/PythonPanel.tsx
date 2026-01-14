import { PyodideRuntime } from "@formula/python-runtime";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

import { ensurePyodideIndexURL } from "../../pyodide/pyodideIndexURL.js";

type NetworkPermission = "none" | "allowlist" | "full";

export type MountPythonPanelOptions = {
  // apps/desktop/src/document/documentController.js
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  documentController: any;
  container: HTMLElement;
  /**
   * Optional hook to keep `formula.active_sheet` aligned with the UI's current sheet.
   */
  getActiveSheetId?: () => string;
};

function parseAllowlist(raw: string) {
  return raw
    .split(/[,\s]+/g)
    .map((host) => host.trim())
    .filter(Boolean);
}

export function mountPythonPanel({ documentController, container, getActiveSheetId }: MountPythonPanelOptions) {
  container.replaceChildren();

  const root = document.createElement("div");
  root.className = "python-panel";

  const toolbar = document.createElement("div");
  toolbar.className = "python-panel__toolbar";

  const runButton = document.createElement("button");
  runButton.type = "button";
  runButton.textContent = "Run";
  runButton.dataset.testid = "python-panel-run";

  const clearButton = document.createElement("button");
  clearButton.type = "button";
  clearButton.textContent = "Clear output";
  clearButton.dataset.testid = "python-clear-output";

  const networkLabel = document.createElement("label");
  networkLabel.textContent = "Network:";
  networkLabel.className = "python-panel__network-label";

  const networkSelect = document.createElement("select");
  networkSelect.dataset.testid = "python-network-permission";
  for (const mode of ["none", "allowlist", "full"] as const) {
    const opt = document.createElement("option");
    opt.value = mode;
    opt.textContent = mode;
    networkSelect.appendChild(opt);
  }
  networkSelect.value = "none";
  networkLabel.appendChild(networkSelect);

  const allowlistInput = document.createElement("input");
  allowlistInput.type = "text";
  allowlistInput.placeholder = "Allowlist hostnames (e.g. example.com api.mycorp.com)";
  allowlistInput.dataset.testid = "python-network-allowlist";
  allowlistInput.className = "python-panel__allowlist-input";

  function updateAllowlistVisibility() {
    allowlistInput.hidden = networkSelect.value !== "allowlist";
  }
  updateAllowlistVisibility();
  networkSelect.addEventListener("change", updateAllowlistVisibility);

  toolbar.appendChild(runButton);
  toolbar.appendChild(clearButton);
  toolbar.appendChild(networkLabel);
  toolbar.appendChild(allowlistInput);

  const split = document.createElement("div");
  split.className = "python-panel__split";

  const editor = document.createElement("textarea");
  editor.dataset.testid = "python-panel-code";
  editor.spellcheck = false;
  editor.value = [
    "import formula",
    "",
    "sheet = formula.active_sheet",
    'sheet[\"A1\"] = 123',
    'print(\"Wrote A1\")',
    "",
  ].join("\n");
  editor.className = "python-panel__editor";

  const output = document.createElement("pre");
  output.dataset.testid = "python-panel-output";
  output.className = "python-panel__output";

  const editorWrap = document.createElement("div");
  editorWrap.className = "python-panel__editor-wrap";
  editorWrap.appendChild(editor);

  split.appendChild(editorWrap);
  split.appendChild(output);

  root.appendChild(toolbar);
  root.appendChild(split);
  container.appendChild(root);

  const bridge = new DocumentControllerBridge(documentController, {
    activeSheetId: getActiveSheetId?.() ?? "Sheet1",
  });

  let disposed = false;
  let initialized = false;

  const runtime = new PyodideRuntime({ api: bridge });

  function effectivePermissions() {
    const network = networkSelect.value as NetworkPermission;
    return {
      filesystem: "none",
      network,
      networkAllowlist: network === "allowlist" ? parseAllowlist(allowlistInput.value) : [],
    };
  }

  clearButton.addEventListener("click", () => {
    output.textContent = "";
  });

  runButton.addEventListener("click", () => {
    void (async () => {
      if (disposed) return;

      // The runtime may reset itself after timeouts/memory errors. Keep our local
      // initialization flag in sync so users can run again without reloading.
      if ((runtime as any).initialized !== true) {
        initialized = false;
      }

      output.textContent =
        runtime.getBackendMode() === "mainThread"
          ? "SharedArrayBuffer unavailable; running Pyodide on main thread (UI may freeze during execution).\n\n"
          : "";
      runButton.disabled = true;

      try {
        const sheetId = getActiveSheetId?.() ?? "Sheet1";
        bridge.activeSheetId = sheetId;
        bridge.sheetIds.add(sheetId);

        const permissions = effectivePermissions();

        if (!initialized) {
          output.textContent += "Preparing Python runtime…\n";
          const indexURL = await ensurePyodideIndexURL({
            onProgress: (progress) => {
              if (progress.kind === "downloadStart" && progress.message) {
                output.textContent += `${progress.message}\n`;
              }
              if (progress.kind === "ready" && progress.message) {
                output.textContent += `${progress.message}\n`;
              }
            },
          });
          output.textContent += "Loading Python runtime…\n";
          await runtime.initialize({ api: bridge, permissions, indexURL });
          initialized = true;
          output.textContent += "Ready.\n\n";
        }

        const result = await runtime.execute(editor.value, { permissions });
        if (typeof result?.stdout === "string" && result.stdout.length > 0) {
          output.textContent += result.stdout;
        }
        if (typeof result?.stderr === "string" && result.stderr.length > 0) {
          output.textContent += result.stderr;
        }
      } catch (err) {
        const error = err as any;
        if (typeof error?.stdout === "string" && error.stdout.length > 0) {
          output.textContent += error.stdout;
        }
        if (typeof error?.stderr === "string" && error.stderr.length > 0) {
          output.textContent += error.stderr;
        }
        output.textContent += `Error: ${error instanceof Error ? error.message : String(error)}\n`;
      } finally {
        runButton.disabled = false;
        output.scrollTop = output.scrollHeight;
      }
    })().catch((err) => {
      // Terminal catch to avoid unhandled rejections from click handlers.
      console.error("Unhandled Python panel error:", err);
      try {
        output.textContent += `Error: ${err instanceof Error ? err.message : String(err)}\n`;
      } catch {
        // ignore output failures
      }
      runButton.disabled = false;
    });
  });

  return () => {
    disposed = true;
    runtime.destroy();
    container.replaceChildren();
  };
}
