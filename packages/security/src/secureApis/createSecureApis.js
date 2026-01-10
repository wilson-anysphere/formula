import fs from "node:fs/promises";

function requirePrincipal(principal) {
  if (!principal || typeof principal.type !== "string" || typeof principal.id !== "string") {
    throw new TypeError("Invalid principal");
  }
}

function requirePermissionManager(permissionManager) {
  if (!permissionManager || typeof permissionManager.ensure !== "function") {
    throw new TypeError("Invalid PermissionManager");
  }
}

export function createSecureFs({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false
}) {
  requirePrincipal(principal);
  requirePermissionManager(permissionManager);

  return {
    async readFile(filePath, options) {
      await permissionManager.ensure(
        principal,
        { kind: "filesystem", access: "read", path: String(filePath) },
        { promptIfDenied }
      );

      auditLogger?.log({
        eventType: "security.filesystem.read",
        actor: principal,
        success: true,
        metadata: { path: String(filePath) }
      });

      return fs.readFile(filePath, options);
    },

    async writeFile(filePath, data, options) {
      await permissionManager.ensure(
        principal,
        { kind: "filesystem", access: "readwrite", path: String(filePath) },
        { promptIfDenied }
      );

      await fs.writeFile(filePath, data, options);

      auditLogger?.log({
        eventType: "security.filesystem.write",
        actor: principal,
        success: true,
        metadata: { path: String(filePath), bytes: Buffer.byteLength(data) }
      });
    }
  };
}

export function createSecureFetch({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false
}) {
  requirePrincipal(principal);
  requirePermissionManager(permissionManager);

  return async function secureFetch(url, init) {
    const urlString = String(url);
    await permissionManager.ensure(principal, { kind: "network", url: urlString }, { promptIfDenied });

    auditLogger?.log({
      eventType: "security.network.request",
      actor: principal,
      success: true,
      metadata: { url: urlString, method: init?.method ?? "GET" }
    });

    return fetch(urlString, init);
  };
}
