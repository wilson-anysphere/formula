/**
 * Script Editor panel scaffold.
 *
 * This repo focuses on the application/controller layer (not a full React UI),
 * but we still keep an explicit module boundary so the desktop shell can mount
 * Monaco and wire the runtime + console output.
 *
 * The actual Monaco dependency is intentionally not included here; the desktop
 * shell can provide it and call `mountScriptEditorPanel`.
 */

import { FORMULA_API_DTS, ScriptRuntime } from "@formula/scripting/web";

/**
 * @typedef {import("@formula/scripting").Workbook} Workbook
 */

/**
 * @param {{
 *   workbook: Workbook,
 *   container: HTMLElement,
 *   monaco?: any,
 * }} params
 * @returns {{ dispose: () => void }}
 */
export function mountScriptEditorPanel({ workbook, container, monaco }) {
  const runtime = new ScriptRuntime(workbook);
  container.innerHTML = "";

  const toolbar = document.createElement("div");
  toolbar.style.display = "flex";
  toolbar.style.gap = "8px";
  toolbar.style.padding = "8px";
  toolbar.style.borderBottom = "1px solid var(--panel-border)";

  const runButton = document.createElement("button");
  runButton.textContent = "Run";
  toolbar.appendChild(runButton);

  const editorHost = document.createElement("div");
  editorHost.style.flex = "1";
  editorHost.style.minHeight = "0";

  const consoleHost = document.createElement("pre");
  consoleHost.style.height = "140px";
  consoleHost.style.margin = "0";
  consoleHost.style.padding = "8px";
  consoleHost.style.overflow = "auto";
  consoleHost.style.borderTop = "1px solid var(--panel-border)";
  consoleHost.textContent = "Output…";

  const root = document.createElement("div");
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.height = "100%";

  root.appendChild(toolbar);
  root.appendChild(editorHost);
  root.appendChild(consoleHost);
  container.appendChild(root);

  let editor = null;
  let currentCode = defaultScript();

  const ensureMonaco = async () => {
    if (editor) return;
    const m = monaco ?? (await import("monaco-editor"));
    m.languages.typescript.typescriptDefaults.addExtraLib(FORMULA_API_DTS, "file:///formula.d.ts");
    editor = m.editor.create(editorHost, {
      value: currentCode,
      language: "typescript",
      automaticLayout: true,
      minimap: { enabled: false },
    });
  };

  const updateConsole = (text) => {
    consoleHost.textContent = text;
  };

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    updateConsole("");
    try {
      await ensureMonaco();
      currentCode = editor.getValue();
      const result = await runtime.run(currentCode);
      const logs = result.logs.map((l) => `[${l.level}] ${l.message}`).join("\n");
      updateConsole(logs + (result.error ? `\n[error] ${result.error.message}` : ""));
    } finally {
      runButton.disabled = false;
    }
  });

  return {
    dispose: () => {
      if (editor) {
        editor.dispose();
        editor = null;
      }
      container.innerHTML = "";
    },
  };
}

function defaultScript() {
  return `export default async function main(ctx: ScriptContext) {
  // Example: read a range, compute sum, write result
  const values = await ctx.activeSheet.getRange("A1:A3").getValues();
  const sum = values.flat().reduce((acc, v) => acc + (typeof v === "number" ? v : 0), 0);

  await ctx.activeSheet.getRange("B1").setValue(sum);
  ctx.ui.log("sum =", sum);
}
`;
}
