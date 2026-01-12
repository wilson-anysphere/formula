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
  toolbar.className = "script-editor__toolbar";

  const runButton = document.createElement("button");
  runButton.textContent = "Run";
  runButton.className = "script-editor__run-button";
  runButton.dataset.testid = "script-editor-run";
  toolbar.appendChild(runButton);

  const editorHost = document.createElement("div");
  editorHost.className = "script-editor__editor-host";

  const consoleHost = document.createElement("pre");
  consoleHost.className = "script-editor__console";
  consoleHost.textContent = "Output…";

  const root = document.createElement("div");
  root.className = "script-editor";

  root.appendChild(toolbar);
  root.appendChild(editorHost);
  root.appendChild(consoleHost);
  container.appendChild(root);

  let editor = null;
  let currentCode = defaultScript();
  const setCodeEvent = "formula:script-editor:set-code";

  const fallbackEditor = document.createElement("textarea");
  fallbackEditor.value = currentCode;
  fallbackEditor.dataset.testid = "script-editor-code";
  fallbackEditor.spellcheck = false;
  fallbackEditor.className = "script-editor__fallback-editor";
  fallbackEditor.addEventListener("input", () => {
    currentCode = fallbackEditor.value;
  });
  editorHost.appendChild(fallbackEditor);

  const handleSetCode = (event) => {
    const next = event?.detail?.code;
    if (typeof next !== "string") return;
    currentCode = next;
    if (editor) {
      editor.setValue(next);
    } else {
      fallbackEditor.value = next;
    }
  };
  window.addEventListener(setCodeEvent, handleSetCode);

  const ensureMonaco = async () => {
    if (editor) return;
    if (!monaco) return;
    const m = monaco;
    m.languages.typescript.typescriptDefaults.addExtraLib(FORMULA_API_DTS, "file:///formula.d.ts");
    editor = m.editor.create(editorHost, {
      value: currentCode,
      language: "typescript",
      automaticLayout: true,
      minimap: { enabled: false },
    });
    fallbackEditor.remove();
  };

  const updateConsole = (text) => {
    consoleHost.textContent = text;
  };

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    updateConsole("");
    try {
      await ensureMonaco();
      currentCode = editor ? editor.getValue() : fallbackEditor.value;
      // Script execution includes worker startup + TypeScript transpilation; use a
      // slightly more forgiving timeout so the first run doesn't flake under load.
      const result = await runtime.run(currentCode, { timeoutMs: 20_000 });
      const logs = result.logs.map((l) => `[${l.level}] ${l.message}`).join("\n");
      updateConsole(logs + (result.error ? `\n[error] ${result.error.message}` : ""));
    } finally {
      runButton.disabled = false;
    }
  });

  return {
    dispose: () => {
      window.removeEventListener(setCodeEvent, handleSetCode);
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
