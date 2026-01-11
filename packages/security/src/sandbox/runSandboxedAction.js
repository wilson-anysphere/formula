import { stripTypeScriptTypes } from "node:module";
import { runSandboxedJavaScript } from "./runSandboxedJavaScript.js";
import { runSandboxedPython } from "./runSandboxedPython.js";

/**
 * Higher-level helper for dispatching sandboxed execution by language/runtime.
 * Preserves the underlying runner interfaces while giving callers a single entrypoint.
 *
 * @param {"javascript" | "js" | "typescript" | "ts" | "python" | "py"} kind
 * @param {any} options
 */
export async function runSandboxedAction(kind, options) {
  if (kind === "python" || kind === "py") {
    return runSandboxedPython(options);
  }

  if (kind === "typescript" || kind === "ts") {
    return runSandboxedJavaScript({
      ...options,
      code: stripTypeScriptTypes(String(options?.code ?? ""))
    });
  }

  if (kind === "javascript" || kind === "js") {
    return runSandboxedJavaScript(options);
  }

  throw new TypeError(`Unsupported sandbox kind: ${String(kind)}`);
}

