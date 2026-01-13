import { describe, expect, it } from "vitest";

import { ContextManager } from "./contextManager.js";

describe("ContextManager range attachments", () => {
  it("includes attached range data in promptContext even when query is unrelated", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      redactor: (text: string) => text,
    });

    const sheet = {
      name: "Sheet1",
      values: [
        ["Name", "Score"],
        ["Alice", 10],
        ["Bob", 20],
      ],
    };

    const out = await cm.buildContext({
      sheet,
      query: "totally unrelated query",
      attachments: [{ type: "range", reference: "A1:B3" }],
    });

    expect(out.promptContext).toContain("## attachment_data");
    expect(out.promptContext).toContain("Sheet1!A1:B3");
    expect(out.promptContext).toContain("Name\tScore");
    expect(out.promptContext).toContain("Alice\t10");
    expect(out.promptContext).toContain("Bob\t20");
  });

  it("translates absolute attachment ranges using sheet.origin", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      redactor: (text: string) => text,
    });

    // This matrix represents K11:L12 (0-index origin row=10,col=10).
    const sheet = {
      name: "Sheet1",
      origin: { row: 10, col: 10 },
      values: [
        ["X", "Y"],
        ["Z", "W"],
      ],
    };

    const out = await cm.buildContext({
      sheet,
      query: "anything",
      attachments: [{ type: "range", reference: "K11:L12" }],
    });

    expect(out.promptContext).toContain("## attachment_data");
    expect(out.promptContext).toContain("Sheet1!K11:L12");
    expect(out.promptContext).toContain("X\tY");
    expect(out.promptContext).toContain("Z\tW");
  });

  it("emits a helpful note (and does not throw) when an attached range is outside the provided sheet window", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      redactor: (text: string) => text,
    });

    const sheet = {
      name: "Sheet1",
      origin: { row: 10, col: 10 },
      values: [["OnlyCell"]],
    };

    const out = await cm.buildContext({
      sheet,
      query: "anything",
      attachments: [{ type: "range", reference: "A1:B2" }],
    });

    expect(out.promptContext).toContain("## attachment_data");
    expect(out.promptContext).toContain("Sheet1!A1:B2");
    expect(out.promptContext).toMatch(/outside the available sheet window/i);
  });
});

