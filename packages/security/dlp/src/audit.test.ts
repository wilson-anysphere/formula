import { describe, expect, it } from "vitest";

import { InMemoryAuditLogger } from "./audit.js";

describe("InMemoryAuditLogger", () => {
  it("marks success=false for BLOCK decisions", () => {
    const audit = new InMemoryAuditLogger();
    audit.log({ type: "ai.cell_function", decision: { decision: "block" } });
    const [event] = audit.list();
    expect(event?.success).toBe(false);
  });

  it("marks success=true for ALLOW/REDACT decisions", () => {
    const audit = new InMemoryAuditLogger();
    audit.log({ type: "ai.cell_function", decision: { decision: "allow" } });
    audit.log({ type: "ai.cell_function", decision: { decision: "redact" } });
    const events = audit.list();
    expect(events[0]?.success).toBe(true);
    expect(events[1]?.success).toBe(true);
  });
});

