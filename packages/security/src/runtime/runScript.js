import { stripTypeScriptTypes } from "node:module";

import { runSandboxedJavaScript } from "../sandbox/runSandboxedJavaScript.js";
import { runSandboxedPython } from "../sandbox/runSandboxedPython.js";

export async function runScript({
  scriptId,
  language = "javascript",
  code,
  permissionManager,
  auditLogger = null,
  timeoutMs,
  memoryMb
}) {
  const principal = { type: "script", id: String(scriptId) };
  const permissionSnapshot = permissionManager.getSnapshot(principal);

  if (language === "python") {
    return runSandboxedPython({
      principal,
      code,
      permissionSnapshot,
      auditLogger,
      timeoutMs,
      memoryMb,
      label: "script"
    });
  }

  let jsCode = code;
  if (language === "typescript" || language === "ts") {
    jsCode = stripTypeScriptTypes(String(code));
  }

  return runSandboxedJavaScript({
    principal,
    code: jsCode,
    permissionSnapshot,
    auditLogger,
    timeoutMs,
    memoryMb,
    label: "script"
  });
}
