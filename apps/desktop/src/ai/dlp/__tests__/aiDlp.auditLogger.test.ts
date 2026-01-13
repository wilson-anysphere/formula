import { beforeEach, describe, expect, it } from "vitest";

import { getAiDlpAuditLogger, resetAiDlpAuditLoggerForTests } from "../aiDlp.js";

describe("Ai DLP audit logger retention", () => {
  beforeEach(() => {
    resetAiDlpAuditLoggerForTests();
  });

  it("caps retained events and drops oldest first (FIFO)", () => {
    const cap = 1_000;
    const logger = getAiDlpAuditLogger();

    const ids: string[] = [];
    for (let i = 0; i < cap + 10; i++) {
      ids.push(logger.log({ type: "test", index: i }));
    }

    const events = logger.list();
    expect(events).toHaveLength(cap);

    const retainedIds = events.map((e: any) => e.id);
    expect(retainedIds).toEqual(ids.slice(-cap));
  });
});

