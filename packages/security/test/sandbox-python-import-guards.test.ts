import { describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runScript } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

describe("Python sandbox import guards", () => {
  it("blocks importing _ctypes", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runScript({
        scriptId: "py.import._ctypes",
        language: "python",
        timeoutMs: 2_000,
        permissionManager,
        auditLogger,
        code: `import _ctypes\n`
      })
    ).rejects.toMatchObject({ code: "PYTHON_SANDBOX_ERROR", name: "ImportError" });
  });

  it("blocks importing _posixsubprocess when automation is not granted", async () => {
    if (process.platform === "win32") return;

    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runScript({
        scriptId: "py.import._posixsubprocess",
        language: "python",
        timeoutMs: 2_000,
        permissionManager,
        auditLogger,
        code: `import _posixsubprocess\n`
      })
    ).rejects.toMatchObject({ code: "PYTHON_SANDBOX_ERROR", name: "ImportError" });
  });
});

