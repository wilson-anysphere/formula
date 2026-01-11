import fs from "node:fs/promises";

import { checkPermissionGrant } from "../permissions.js";

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

function createUnimplementedError(message) {
  const err = new Error(message);
  err.code = "SECURE_API_UNAVAILABLE";
  return err;
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

export function createSecureClipboard({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false,
  adapter = null
}) {
  requirePrincipal(principal);
  requirePermissionManager(permissionManager);

  return {
    async readText() {
      await permissionManager.ensure(principal, { kind: "clipboard" }, { promptIfDenied });
      if (!adapter?.readText) throw createUnimplementedError("Clipboard adapter readText() not configured");

      const text = await adapter.readText();
      auditLogger?.log({
        eventType: "security.clipboard.read",
        actor: principal,
        success: true
      });
      return text;
    },
    async writeText(text) {
      await permissionManager.ensure(principal, { kind: "clipboard" }, { promptIfDenied });
      if (!adapter?.writeText) throw createUnimplementedError("Clipboard adapter writeText() not configured");

      await adapter.writeText(String(text));
      auditLogger?.log({
        eventType: "security.clipboard.write",
        actor: principal,
        success: true,
        metadata: { bytes: Buffer.byteLength(String(text)) }
      });
    }
  };
}

export function createSecureNotifications({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false,
  adapter = null
}) {
  requirePrincipal(principal);
  requirePermissionManager(permissionManager);

  return {
    async notify(notification) {
      await permissionManager.ensure(principal, { kind: "notifications" }, { promptIfDenied });
      if (!adapter?.notify) throw createUnimplementedError("Notifications adapter notify() not configured");

      await adapter.notify(notification);
      auditLogger?.log({
        eventType: "security.notifications.notify",
        actor: principal,
        success: true
      });
    }
  };
}

export function createSecureAutomation({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false,
  adapter = null
}) {
  requirePrincipal(principal);
  requirePermissionManager(permissionManager);

  return {
    async run(action) {
      await permissionManager.ensure(principal, { kind: "automation" }, { promptIfDenied });
      if (!adapter?.run) throw createUnimplementedError("Automation adapter run() not configured");

      const result = await adapter.run(action);
      auditLogger?.log({
        eventType: "security.automation.run",
        actor: principal,
        success: true,
        metadata: { action: action?.type ?? null }
      });
      return result;
    }
  };
}

export function createSecureApis({
  principal,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false,
  adapters = {}
}) {
  return {
    fs: createSecureFs({ principal, permissionManager, auditLogger, promptIfDenied }),
    fetch: createSecureFetch({ principal, permissionManager, auditLogger, promptIfDenied }),
    clipboard: createSecureClipboard({
      principal,
      permissionManager,
      auditLogger,
      promptIfDenied,
      adapter: adapters.clipboard ?? null
    }),
    notifications: createSecureNotifications({
      principal,
      permissionManager,
      auditLogger,
      promptIfDenied,
      adapter: adapters.notifications ?? null
    }),
    automation: createSecureAutomation({
      principal,
      permissionManager,
      auditLogger,
      promptIfDenied,
      adapter: adapters.automation ?? null
    })
  };
}

function createSandboxError(payload) {
  const err = Object.create(null);
  if (payload && typeof payload === "object") {
    for (const [key, value] of Object.entries(payload)) {
      err[key] = value;
    }
  }
  return err;
}

function normalizeSandboxError(error) {
  if (error && typeof error === "object") {
    return createSandboxError({
      name: typeof error.name === "string" ? error.name : "Error",
      message: typeof error.message === "string" ? error.message : String(error),
      code: typeof error.code === "string" ? error.code : undefined,
      stack: typeof error.stack === "string" ? error.stack : undefined
    });
  }
  return createSandboxError({ name: "Error", message: String(error) });
}

function hardenFunction(fn) {
  if (typeof fn !== "function") return fn;
  Object.setPrototypeOf(fn, null);
  return fn;
}

function hardenObject(obj) {
  if (!obj || typeof obj !== "object") return obj;
  Object.setPrototypeOf(obj, null);
  return Object.freeze(obj);
}

function createSandboxPermissionEnsurer({ principal, permissionSnapshot, auditLogger }) {
  const snapshot = permissionSnapshot ?? null;

  return async (request) => {
    const decision = checkPermissionGrant(snapshot, request);

    auditLogger?.log({
      eventType: "security.permission.checked",
      actor: principal,
      success: decision.allowed,
      metadata: {
        request,
        ...(decision.allowed ? {} : { reason: decision.reason })
      }
    });

    if (decision.allowed) return;

    auditLogger?.log({
      eventType: "security.permission.denied",
      actor: principal,
      success: false,
      metadata: { request, reason: decision.reason }
    });

    throw createSandboxError({
      name: "PermissionDeniedError",
      message: decision.reason,
      code: "PERMISSION_DENIED",
      principal,
      request,
      reason: decision.reason
    });
  };
}

export function createSandboxSecureApis({
  principal,
  permissionSnapshot,
  auditLogger = null
}) {
  requirePrincipal(principal);

  const ensure = createSandboxPermissionEnsurer({ principal, permissionSnapshot, auditLogger });

  const sandboxFs = hardenObject({
    readFile: hardenFunction(async (filePath, options) => {
      const pathString = String(filePath);
      await ensure({ kind: "filesystem", access: "read", path: pathString });

      auditLogger?.log({
        eventType: "security.filesystem.read",
        actor: principal,
        success: true,
        metadata: { path: pathString }
      });

      let encoding = "utf8";
      if (typeof options === "string") encoding = options;
      if (options && typeof options === "object" && typeof options.encoding === "string") {
        encoding = options.encoding;
      }

      try {
        return await fs.readFile(pathString, { encoding });
      } catch (error) {
        throw normalizeSandboxError(error);
      }
    }),
    writeFile: hardenFunction(async (filePath, data, options) => {
      const pathString = String(filePath);
      await ensure({ kind: "filesystem", access: "readwrite", path: pathString });

      const content = typeof data === "string" ? data : String(data);
      try {
        await fs.writeFile(pathString, content, options);
      } catch (error) {
        throw normalizeSandboxError(error);
      }

      auditLogger?.log({
        eventType: "security.filesystem.write",
        actor: principal,
        success: true,
        metadata: { path: pathString, bytes: Buffer.byteLength(content) }
      });
    })
  });

  const sandboxFetch = hardenFunction(async (url, init) => {
    const urlString = String(url);
    await ensure({ kind: "network", url: urlString });

    auditLogger?.log({
      eventType: "security.network.request",
      actor: principal,
      success: true,
      metadata: { url: urlString, method: init?.method ?? "GET" }
    });

    let res;
    try {
      res = await fetch(urlString, init);
    } catch (error) {
      throw normalizeSandboxError(error);
    }

    const headers = Object.create(null);
    for (const [key, value] of res.headers.entries()) {
      headers[key] = value;
    }

    const response = Object.create(null);
    response.ok = Boolean(res.ok);
    response.status = Number(res.status);
    response.url = String(res.url);
    response.headers = hardenObject(headers);
    response.text = hardenFunction(async () => {
      try {
        return await res.text();
      } catch (error) {
        throw normalizeSandboxError(error);
      }
    });
    return hardenObject(response);
  });

  const sandboxClipboard = hardenObject({
    readText: hardenFunction(async () => {
      await ensure({ kind: "clipboard" });
      throw createSandboxError({
        name: "Error",
        message: "Clipboard API is not available in this sandbox",
        code: "SECURE_API_UNAVAILABLE"
      });
    }),
    writeText: hardenFunction(async () => {
      await ensure({ kind: "clipboard" });
      throw createSandboxError({
        name: "Error",
        message: "Clipboard API is not available in this sandbox",
        code: "SECURE_API_UNAVAILABLE"
      });
    })
  });

  const sandboxNotifications = hardenObject({
    notify: hardenFunction(async () => {
      await ensure({ kind: "notifications" });
      throw createSandboxError({
        name: "Error",
        message: "Notifications API is not available in this sandbox",
        code: "SECURE_API_UNAVAILABLE"
      });
    })
  });

  const sandboxAutomation = hardenObject({
    run: hardenFunction(async () => {
      await ensure({ kind: "automation" });
      throw createSandboxError({
        name: "Error",
        message: "Automation API is not available in this sandbox",
        code: "SECURE_API_UNAVAILABLE"
      });
    })
  });

  const apis = Object.create(null);
  apis.fs = sandboxFs;
  apis.fetch = sandboxFetch;
  apis.clipboard = sandboxClipboard;
  apis.notifications = sandboxNotifications;
  apis.automation = sandboxAutomation;

  return hardenObject(apis);
}
