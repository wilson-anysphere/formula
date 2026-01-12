import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runExtension } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

describe("Sandbox audit trail", () => {
  it("emits start -> permission check -> operation -> complete for allowed operations", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-sec-audit-"));
    const filePath = path.join(dir, "data.txt");
    await fs.writeFile(filePath, "audit", "utf8");

    const { auditLogger, events } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const extensionId = "ext.audit.sequence";
    const principal = { type: "extension", id: extensionId };
    permissionManager.grant(principal, { filesystem: { read: [dir] } });

    // Ignore grant auditing; we only care about a single run's events.
    events.length = 0;

    await runExtension({
      extensionId,
      permissionManager,
      auditLogger,
      timeoutMs: 5_000,
      code: `await fs.readFile(${JSON.stringify(filePath)}, "utf8");\nreturn null;`
    });

    // AuditLogger normalizes the legacy `metadata` field into `details` on the
    // canonical AuditEvent schema.
    const types = events.map((e) => [e.eventType, e.details?.phase ?? null]);

    // Ensure ordering within the run.
    expect(types[0]).toEqual(["security.extension.run", "start"]);
    expect(types.some(([t]) => t === "security.permission.checked")).toBe(true);
    expect(types.some(([t]) => t === "security.filesystem.read")).toBe(true);
    expect(types[types.length - 1]).toEqual(["security.extension.run", "complete"]);

    const permIdx = types.findIndex(([t]) => t === "security.permission.checked");
    const opIdx = types.findIndex(([t]) => t === "security.filesystem.read");
    expect(permIdx).toBeGreaterThan(0);
    expect(opIdx).toBeGreaterThan(permIdx);
  });
});
