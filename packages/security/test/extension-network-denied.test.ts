import { describe, expect, it } from "vitest";

import {
  AuditLogger,
  PermissionManager,
  SqliteAuditLogStore,
  runExtension
} from "../src/index.js";

describe("Extension sandbox", () => {
  it("denies network access by default without prompting", async () => {
    const store = new SqliteAuditLogStore({ path: ":memory:" });
    const auditLogger = new AuditLogger({ store });

    let promptCalls = 0;
    const permissionManager = new PermissionManager({
      auditLogger,
      onPrompt: async () => {
        promptCalls += 1;
        return true;
      }
    });

    await expect(
      runExtension({
        extensionId: "ext.network",
        permissionManager,
        auditLogger,
        timeoutMs: 2_000,
        code: `await fetch("https://example.com")`
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "network" } });

    expect(promptCalls).toBe(0);

    const deniedEvents = store.query({ eventType: "security.permission.denied" });
    expect(
      deniedEvents.some((e) => e.actor.type === "extension" && e.actor.id === "ext.network")
    ).toBe(true);
  });
});

