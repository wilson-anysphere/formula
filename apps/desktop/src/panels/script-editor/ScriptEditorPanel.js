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
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../../collab/permissionGuards.js";

/**
 * @typedef {import("@formula/scripting").Workbook} Workbook
 */

/**
 * @param {{
 *   workbook: Workbook,
 *   container: HTMLElement,
 *   monaco?: any,
 *   isEditing?: () => boolean,
 *   isReadOnly?: () => boolean,
 * }} params
 * @returns {{ dispose: () => void }}
 */
export function mountScriptEditorPanel({ workbook, container, monaco, isEditing, isReadOnly }) {
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
  const abort = typeof AbortController !== "undefined" ? new AbortController() : null;
  const eventSignal = abort?.signal;
  let isRunning = false;
  let lastEditing = (globalThis.__formulaSpreadsheetIsEditing ?? false) === true;
  let lastReadOnly = false;
  let blockedOutputReason = null;

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
    if (getEditingState()) return "Finish editing to run scripts.";
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
        updateConsole(reason);
        blockedOutputReason = reason;
      }
      return;
    }
    runButton.removeAttribute("title");
    delete runButton.dataset.blockedReason;
    const current = String(consoleHost.textContent ?? "").trim();
    if (blockedOutputReason && current === blockedOutputReason.trim()) {
      updateConsole("Output…");
    }
    blockedOutputReason = null;
  };

  syncRunButtonDisabledState();

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

  runButton.addEventListener("click", () => {
    void (async () => {
      const blockedReason = getBlockedReason();
      if (blockedReason) {
        updateConsole(blockedReason);
        try {
          syncRunButtonDisabledState();
        } catch {
          // ignore UI sync failures
        }
        return;
      }

      isRunning = true;
      syncRunButtonDisabledState();
      updateConsole("");
      try {
        await ensureMonaco();
        currentCode = editor ? editor.getValue() : fallbackEditor.value;
        // Script execution includes worker startup + TypeScript transpilation; use a
        // slightly more forgiving timeout so the first run doesn't flake under load.
        const result = await runtime.run(currentCode, { timeoutMs: 20_000 });
        const logs = result.logs.map((l) => `[${l.level}] ${l.message}`).join("\n");
        updateConsole(logs + (result.error ? `\n[error] ${result.error.message}` : ""));
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        updateConsole(`[error] ${message}`);
      } finally {
        isRunning = false;
        try {
          syncRunButtonDisabledState();
        } catch {
          // ignore UI sync failures
        }
      }
    })().catch((err) => {
      // `click` handlers ignore returned promises, so ensure we always attach a
      // terminal catch to avoid unhandled rejections.
      try {
        const message = err instanceof Error ? err.message : String(err);
        updateConsole(`[error] ${message}`);
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
