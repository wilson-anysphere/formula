import { describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runExtension } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

// These sandbox tests can be CPU/IO sensitive under heavily parallelized CI shards.
// Keep the timeout comfortably above typical cold-start overhead so we don't flake.
const SANDBOX_TIMEOUT_MS = 30_000;

describe("Sandbox optional capability enforcement", () => {
  it("denies clipboard/notifications/automation by default in JS sandbox", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runExtension({
        extensionId: "ext.clipboard.denied",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.clipboard.writeText("hello");`
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "clipboard" } });

    await expect(
      runExtension({
        extensionId: "ext.notifications.denied",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.notifications.notify({ title: "hi" });`
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "notifications" } });

    await expect(
      runExtension({
        extensionId: "ext.automation.denied",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.automation.run({ type: "noop" });`
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "automation" } });
  });

  it("still gates access even when the API is unavailable in the sandbox runtime", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const clipboardPrincipal = { type: "extension", id: "ext.clipboard.allowed" };
    permissionManager.grant(clipboardPrincipal, { clipboard: true });

    await expect(
      runExtension({
        extensionId: "ext.clipboard.allowed",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.clipboard.writeText("hello");`
      })
    ).rejects.toMatchObject({ code: "SECURE_API_UNAVAILABLE" });

    const notificationsPrincipal = { type: "extension", id: "ext.notifications.allowed" };
    permissionManager.grant(notificationsPrincipal, { notifications: true });

    await expect(
      runExtension({
        extensionId: "ext.notifications.allowed",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.notifications.notify({ title: "hi" });`
      })
    ).rejects.toMatchObject({ code: "SECURE_API_UNAVAILABLE" });

    const automationPrincipal = { type: "extension", id: "ext.automation.allowed" };
    permissionManager.grant(automationPrincipal, { automation: true });

    await expect(
      runExtension({
        extensionId: "ext.automation.allowed",
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS,
        code: `await SecureApis.automation.run({ type: "noop" });`
      })
    ).rejects.toMatchObject({ code: "SECURE_API_UNAVAILABLE" });
  });
});
