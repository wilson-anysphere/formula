import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import {
  AuditLogger,
  PermissionManager,
  SqliteAuditLogStore,
  createSecureFs
} from "../src/index.js";

describe("Security permissions", () => {
  it("denies filesystem writes by default for AI tool calls", async () => {
    const store = new SqliteAuditLogStore({ path: ":memory:" });
    const auditLogger = new AuditLogger({ store });
    const permissionManager = new PermissionManager({ auditLogger });

    const principal = { type: "ai", id: "session-1" };
    const secureFs = createSecureFs({ principal, permissionManager, auditLogger });

    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-sec-"));
    const filePath = path.join(dir, "out.txt");

    await expect(secureFs.writeFile(filePath, "hello")).rejects.toMatchObject({
      code: "PERMISSION_DENIED",
      request: { kind: "filesystem", access: "readwrite" }
    });

    await assert.rejects(() => fs.stat(filePath));

    const deniedEvents = store.query({ eventType: "security.permission.denied" });
    expect(deniedEvents.some((e) => e.actor.type === "ai" && e.actor.id === "session-1")).toBe(true);
  });
});

