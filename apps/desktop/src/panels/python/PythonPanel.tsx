import { PyodideRuntime } from "@formula/python-runtime";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

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
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.height = "100%";
  root.style.padding = "8px";
  root.style.boxSizing = "border-box";
  root.style.gap = "8px";
  root.style.fontSize = "12px";

  const toolbar = document.createElement("div");
  toolbar.style.display = "flex";
  toolbar.style.flexWrap = "wrap";
  toolbar.style.alignItems = "center";
  toolbar.style.gap = "8px";

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
  networkLabel.style.display = "inline-flex";
  networkLabel.style.alignItems = "center";
  networkLabel.style.gap = "6px";

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
  allowlistInput.style.minWidth = "280px";

  function updateAllowlistVisibility() {
    allowlistInput.style.display = networkSelect.value === "allowlist" ? "inline-block" : "none";
  }
  updateAllowlistVisibility();
  networkSelect.addEventListener("change", updateAllowlistVisibility);

  toolbar.appendChild(runButton);
  toolbar.appendChild(clearButton);
  toolbar.appendChild(networkLabel);
  toolbar.appendChild(allowlistInput);

  const split = document.createElement("div");
  split.style.display = "grid";
  split.style.gridTemplateRows = "1fr 1fr";
  split.style.gap = "8px";
  split.style.flex = "1";
  split.style.minHeight = "0";

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
  editor.style.width = "100%";
  editor.style.height = "100%";
  editor.style.resize = "none";
  editor.style.fontFamily =
    'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace';
  editor.style.fontSize = "12px";
  editor.style.lineHeight = "16px";
  editor.style.padding = "8px";
  editor.style.boxSizing = "border-box";

  const output = document.createElement("pre");
  output.dataset.testid = "python-panel-output";
  output.style.margin = "0";
  output.style.width = "100%";
  output.style.height = "100%";
  output.style.overflow = "auto";
  output.style.whiteSpace = "pre-wrap";
  output.style.fontFamily =
    'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace';
  output.style.fontSize = "12px";
  output.style.lineHeight = "16px";
  output.style.padding = "8px";
  output.style.boxSizing = "border-box";
  output.style.background = "var(--bg-secondary)";
  output.style.border = "1px solid var(--border)";
  output.style.borderRadius = "6px";

  const editorWrap = document.createElement("div");
  editorWrap.style.border = "1px solid var(--border)";
  editorWrap.style.borderRadius = "6px";
  editorWrap.style.overflow = "hidden";
  editorWrap.style.minHeight = "0";
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

  runButton.addEventListener("click", async () => {
    if (disposed) return;

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
        output.textContent += "Loading Python runtime…\n";
        await runtime.initialize({ api: bridge, permissions });
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
  });

  return () => {
    disposed = true;
    runtime.destroy();
    container.replaceChildren();
  };
}
