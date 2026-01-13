import { describe, expect, it, vi } from "vitest";

import { ContextManager } from "./contextManager.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

const POLICY_REDACT_RESTRICTED = {
  version: 1,
  allowDocumentOverrides: true,
  rules: {
    [DLP_ACTION.AI_CLOUD_PROCESSING]: {
      maxAllowed: "Internal",
      allowRestrictedContent: false,
      redactDisallowed: true,
    },
  },
} as const;

describe("ContextManager.buildContext origin-aware handling", () => {
  it("redacts classified cells using absolute coordinates when sheet.origin is set", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000,
      // Disable heuristic redaction so assertions are deterministic.
      redactor: (text: string) => text,
    });

    const auditLogger = { log: vi.fn() };

    const out = await cm.buildContext({
      sheet: {
        name: "Sheet1",
        origin: { row: 10, col: 10 },
        values: [
          ["a", "b"],
          ["c", "TOP SECRET"],
        ],
      },
      query: "secret",
      dlp: {
        documentId: "doc-1",
        sheetId: "Sheet1",
        policy: POLICY_REDACT_RESTRICTED,
        classificationRecords: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 11, col: 11 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
        auditLogger,
      },
    });

    expect(JSON.stringify(out.sampledRows)).not.toContain("TOP SECRET");
    expect(out.sampledRows[1]?.[1]).toBe("[REDACTED]");
    expect(out.promptContext).toContain("[REDACTED]");

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
  });

  it("does not let classified cells outside the origin window affect redaction", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000,
      redactor: (text: string) => text,
    });

    const out = await cm.buildContext({
      sheet: {
        name: "Sheet1",
        origin: { row: 10, col: 10 },
        values: [
          ["a", "b"],
          ["c", "d"],
        ],
      },
      query: "a",
      dlp: {
        documentId: "doc-1",
        sheetId: "Sheet1",
        policy: POLICY_REDACT_RESTRICTED,
        // Restricted cell at an absolute coordinate outside the provided window.
        classificationRecords: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    expect(out.sampledRows[0]?.[1]).toBe("b");
    expect(out.sampledRows[1]?.[1]).toBe("d");
    expect(out.promptContext).not.toContain("[REDACTED]");
  });

  it("returns origin-aware retrieved ranges and chunk previews", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000,
      redactor: (text: string) => text,
    });

    const sheet = {
      name: "Sheet1",
      origin: { row: 10, col: 5 },
      values: [
        ["Region", "Revenue"],
        ["North", 1000],
        ["South", 2000],
      ],
    };

    const out = await cm.buildContext({ sheet, query: "revenue by region" });

    expect(out.schema.dataRegions[0]?.range).toBe("Sheet1!F11:G13");
    expect(out.retrieved[0]?.range).toBe("Sheet1!F11:G13");
    expect(out.retrieved[0]?.preview).toContain("Revenue");
  });

  it("redacts structured return fields when DLP heuristic redaction is active", async () => {
    const cm = new ContextManager({ tokenBudgetTokens: 1_000 });

    const auditLogger = { log: vi.fn() };

    const out = await cm.buildContext({
      sheet: {
        name: "Sheet1",
        values: [
          ["Name", "Email"],
          ["Alice", "alice@example.com"],
        ],
      },
      query: "alice",
      dlp: {
        documentId: "doc-1",
        sheetId: "Sheet1",
        policy: POLICY_REDACT_RESTRICTED,
        classificationRecords: [],
        auditLogger,
      },
    });

    expect(out.promptContext).toContain("DLP:");
    expect(out.promptContext).not.toContain("alice@example.com");

    // Prompt-safe structured returns: do not leak sensitive values.
    expect(JSON.stringify(out.sampledRows)).not.toContain("alice@example.com");
    expect(JSON.stringify(out.schema)).not.toContain("alice@example.com");
    expect(JSON.stringify(out.retrieved)).not.toContain("alice@example.com");

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
  });

  it("enforces heuristic redaction under DLP even when the configured redactor is a no-op", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000,
      redactor: (text: string) => text,
    });

    const auditLogger = { log: vi.fn() };

    const out = await cm.buildContext({
      sheet: {
        name: "Sheet1",
        values: [
          ["Name", "Email"],
          ["Alice", "alice@example.com"],
        ],
      },
      query: "alice",
      dlp: {
        documentId: "doc-1",
        sheetId: "Sheet1",
        policy: POLICY_REDACT_RESTRICTED,
        classificationRecords: [],
        auditLogger,
      },
    });

    expect(out.promptContext).toContain("DLP:");
    expect(out.promptContext).toContain("[REDACTED]");
    expect(out.promptContext).not.toContain("alice@example.com");

    expect(JSON.stringify(out.sampledRows)).not.toContain("alice@example.com");
    expect(JSON.stringify(out.schema)).not.toContain("alice@example.com");
    expect(JSON.stringify(out.retrieved)).not.toContain("alice@example.com");

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
  });
});
