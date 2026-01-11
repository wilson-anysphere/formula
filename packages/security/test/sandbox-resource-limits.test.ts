import { describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runExtension, runSandboxedAction, runScript } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

describe("Sandbox resource limits", () => {
  it("enforces JavaScript timeouts", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runExtension({
        extensionId: "ext.timeout",
        code: `while (true) {}`,
        permissionManager,
        auditLogger,
        timeoutMs: 200
      })
    ).rejects.toMatchObject({ code: "SANDBOX_TIMEOUT" });
  });

  it("enforces Python timeouts", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runScript({
        scriptId: "py.timeout",
        language: "python",
        code: `while True:\n    pass\n`,
        permissionManager,
        auditLogger,
        timeoutMs: 200
      })
    ).rejects.toMatchObject({ code: "SANDBOX_TIMEOUT" });
  });

  it("terminates JavaScript workers under memory pressure (best-effort)", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runExtension({
        extensionId: "ext.memory",
        permissionManager,
        auditLogger,
        timeoutMs: 10_000,
        memoryMb: 32,
        code: `
          const blobs = [];
          while (true) {
            blobs.push("x".repeat(1024));
            if (blobs.length % 1000 === 0) {
              await new Promise((r) => setTimeout(r, 0));
            }
          }
        `
      })
    ).rejects.toMatchObject({ code: "SANDBOX_MEMORY_LIMIT" });
  });

  it("enforces JavaScript output limits", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const principal = { type: "extension", id: "ext.output" };
    const permissionSnapshot = permissionManager.getSnapshot(principal);

    await expect(
      runSandboxedAction("javascript", {
        principal,
        permissionSnapshot,
        auditLogger,
        timeoutMs: 2_000,
        maxOutputBytes: 1024,
        code: `
          for (let i = 0; i < 50; i++) {
            console.log("x".repeat(200));
          }
        `
      })
    ).rejects.toMatchObject({ code: "SANDBOX_OUTPUT_LIMIT" });
  });
});
