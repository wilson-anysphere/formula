import fs from "node:fs/promises";
import vm from "node:vm";
import { parentPort } from "node:worker_threads";

import { PermissionDeniedError } from "../errors.js";
import { isPathWithinScope, isUrlAllowedByAllowlist, normalizeScopePath } from "../permissions.js";

function serializeError(error) {
  if (!error || typeof error !== "object") {
    return { name: "Error", message: String(error) };
  }

  return {
    name: error.name,
    message: error.message,
    code: error.code,
    stack: error.stack,
    principal: error.principal,
    request: error.request,
    reason: error.reason
  };
}

function sendAudit(principal, eventType, success, metadata) {
  parentPort.postMessage({
    type: "audit",
    event: { eventType, actor: principal, success: Boolean(success), metadata: metadata ?? {} }
  });
}

function assertNetworkAllowed({ permissions, principal, url }) {
  const mode = permissions?.network?.mode ?? "none";
  if (mode === "full") return;

  if (mode === "allowlist") {
    const allowlist = permissions?.network?.allowlist ?? [];
    if (isUrlAllowedByAllowlist(url, allowlist)) return;
  }

  const request = { kind: "network", url };
  sendAudit(principal, "security.permission.denied", false, { request });
  throw new PermissionDeniedError({
    principal,
    request,
    reason: `Network access denied for ${url}`
  });
}

function resolvePath(p) {
  return normalizeScopePath(p);
}

function assertFilesystemAllowed({ permissions, principal, access, filePath }) {
  const absPath = resolvePath(filePath);
  const readScopes = [
    ...(permissions?.filesystem?.read ?? []),
    ...(permissions?.filesystem?.readwrite ?? [])
  ].map(resolvePath);

  const writeScopes = (permissions?.filesystem?.readwrite ?? []).map(resolvePath);

  if (access === "readwrite") {
    for (const scope of writeScopes) {
      if (isPathWithinScope(absPath, scope)) return;
    }
  } else {
    for (const scope of readScopes) {
      if (isPathWithinScope(absPath, scope)) return;
    }
  }

  const request = { kind: "filesystem", access, path: absPath };
  sendAudit(principal, "security.permission.denied", false, { request });
  throw new PermissionDeniedError({
    principal,
    request,
    reason: `Filesystem ${access} access denied for ${absPath}`
  });
}

parentPort.on("message", async (message) => {
  if (!message || typeof message !== "object" || message.type !== "run") return;

  const { principal, permissions, code, timeoutMs } = message;

  const sandboxFetch = async (url, init) => {
    const urlString = String(url);
    assertNetworkAllowed({ permissions, principal, url: urlString });
    sendAudit(principal, "security.network.request", true, { url: urlString, method: init?.method ?? "GET" });
    return fetch(urlString, init);
  };

  const sandboxFs = {
    readFile: async (filePath, options) => {
      assertFilesystemAllowed({ permissions, principal, access: "read", filePath: String(filePath) });
      sendAudit(principal, "security.filesystem.read", true, { path: String(filePath) });
      return fs.readFile(filePath, options);
    },
    writeFile: async (filePath, data, options) => {
      assertFilesystemAllowed({
        permissions,
        principal,
        access: "readwrite",
        filePath: String(filePath)
      });
      await fs.writeFile(filePath, data, options);
      sendAudit(principal, "security.filesystem.write", true, { path: String(filePath) });
    }
  };

  const sandbox = {
    console,
    fetch: sandboxFetch,
    fs: sandboxFs,
    URL,
    TextEncoder,
    TextDecoder,
    setTimeout,
    clearTimeout
  };

  const context = vm.createContext(sandbox, {
    name: "formula-sandbox",
    codeGeneration: { strings: false, wasm: false }
  });

  try {
    const wrapped = `(async () => {\n${code}\n})()`;
    const script = new vm.Script(wrapped, { filename: "sandboxed-script.js" });

    // vm's timeout only applies to the initial synchronous execution of the script.
    // The parent thread additionally enforces an overall timeout and will terminate
    // the worker if the promise does not settle.
    const result = await script.runInContext(context, { timeout: timeoutMs });
    parentPort.postMessage({ type: "result", result });
  } catch (error) {
    parentPort.postMessage({ type: "error", error: serializeError(error) });
  }
});
