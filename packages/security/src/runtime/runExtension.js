import { runSandboxedJavaScript } from "../sandbox/runSandboxedJavaScript.js";

export async function runExtension({
  extensionId,
  code,
  permissionManager,
  auditLogger = null,
  timeoutMs,
  memoryMb
}) {
  const principal = { type: "extension", id: String(extensionId) };
  const permissionSnapshot = permissionManager.getSnapshot(principal);

  return runSandboxedJavaScript({
    principal,
    code,
    permissionSnapshot,
    auditLogger,
    timeoutMs,
    memoryMb,
    label: "extension"
  });
}
