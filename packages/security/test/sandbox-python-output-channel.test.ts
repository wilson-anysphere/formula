import { describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runScript } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

describe("Python sandbox output channel", () => {
  it("captures writes to sys.__stdout__ and os.write without corrupting JSON output", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    await expect(
      runScript({
        scriptId: "py.output.channel",
        language: "python",
        timeoutMs: 2_000,
        permissionManager,
        auditLogger,
        code: `
import os
import sys

sys.__stdout__.write("A\\\\n")
os.write(1, b"B\\\\n")

fd = os.dup(1)
os.write(fd, b"C\\\\n")

__result__ = 7
`.trim()
      })
    ).resolves.toBe(7);
  });
});

