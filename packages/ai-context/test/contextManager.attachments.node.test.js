import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

test("ContextManager range attachments: includes attached range data in promptContext even when query is unrelated", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
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

  assert.ok(out.promptContext.includes("## attachment_data"));
  assert.ok(out.promptContext.includes("Sheet1!A1:B3"));
  assert.ok(out.promptContext.includes("Name\tScore"));
  assert.ok(out.promptContext.includes("Alice\t10"));
  assert.ok(out.promptContext.includes("Bob\t20"));
});

test("ContextManager range attachments: translates absolute attachment ranges using sheet.origin", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
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

  assert.ok(out.promptContext.includes("## attachment_data"));
  assert.ok(out.promptContext.includes("Sheet1!K11:L12"));
  assert.ok(out.promptContext.includes("X\tY"));
  assert.ok(out.promptContext.includes("Z\tW"));
});

test("ContextManager range attachments: emits a helpful note (and does not throw) when an attached range is outside the provided sheet window", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
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

  assert.ok(out.promptContext.includes("## attachment_data"));
  assert.ok(out.promptContext.includes("Sheet1!A1:B2"));
  assert.match(out.promptContext, /outside the available sheet window/i);
});

