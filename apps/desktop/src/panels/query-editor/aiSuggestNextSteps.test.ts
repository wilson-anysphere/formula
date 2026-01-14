import { beforeEach, describe, expect, it, vi } from "vitest";

import { DataTable, type Query } from "@formula/power-query";

const chatMock = vi.fn();

vi.mock("../../ai/llm/desktopLLMClient.js", () => ({
  getDesktopLLMClient: () => ({ chat: chatMock }),
  getDesktopModel: () => "gpt-4o-mini",
}));

const { suggestQueryNextSteps } = await import("./aiSuggestNextSteps.js");

function baseQuery(): Query {
  return { id: "q1", name: "Query 1", source: { type: "range", range: { values: [] } }, steps: [] };
}

describe("suggestQueryNextSteps", () => {
  beforeEach(() => {
    chatMock.mockReset();
  });

  it("parses a JSON array of operations and validates against schema columns", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
          { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] },
        ]),
      },
    });

    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Sales", type: "number" },
      ],
      [],
    );

    const ops = await suggestQueryNextSteps("filter and sort", { query: baseQuery(), preview });
    expect(ops).toEqual([
      { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] },
    ]);
  });

  it("drops operations that reference unknown columns (keeps valid ones)", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "filterRows", predicate: { type: "comparison", column: "DOES_NOT_EXIST", operator: "isNotNull" } },
          { type: "take", count: 10 },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("do something", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 10 }]);
  });

  it("supports code-fenced JSON responses", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: "```json\n[{\"type\":\"take\",\"count\":5}]\n```",
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("keep top rows", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 5 }]);
  });

  it("accepts a single operation object response", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify({ type: "take", count: 3 }),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("keep top rows", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 3 }]);
  });

  it("limits the number of returned operations to 3", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "take", count: 1 },
          { type: "take", count: 2 },
          { type: "take", count: 3 },
          { type: "take", count: 4 },
          { type: "take", count: 5 },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("keep top rows", { query: baseQuery(), preview });
    expect(ops).toEqual([
      { type: "take", count: 1 },
      { type: "take", count: 2 },
      { type: "take", count: 3 },
    ]);
  });

  it("when schema is missing, only allows schema-independent operations", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } },
          { type: "addColumn", name: "Flag", formula: "1" },
          { type: "removeColumns", columns: [] },
          { type: "take", count: 25 },
          { type: "distinctRows", columns: null },
          { type: "removeRowsWithErrors", columns: null },
        ]),
      },
    });

    const ops = await suggestQueryNextSteps("remove duplicates", { query: baseQuery(), preview: null });
    expect(ops).toEqual([
      { type: "take", count: 25 },
      { type: "distinctRows", columns: null },
      { type: "removeRowsWithErrors", columns: null },
    ]);
  });

  it("throws a helpful error when the model returns invalid JSON", async () => {
    chatMock.mockResolvedValue({
      message: { role: "assistant", content: "not json" },
    });

    await expect(suggestQueryNextSteps("x", { query: baseQuery(), preview: null })).rejects.toThrow(
      /AI returned invalid JSON/i,
    );
  });
});
