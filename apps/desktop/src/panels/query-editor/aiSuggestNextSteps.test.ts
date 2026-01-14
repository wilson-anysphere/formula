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

  it("drops filterRows comparisons that are missing required values", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals" } },
          { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "contains", value: null } },
          { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("filter", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } }]);
  });

  it("drops take operations with invalid counts", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "take", count: -1 },
          { type: "take", count: 3.5 },
          { type: "take", count: 7 },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("top rows", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 7 }]);
  });

  it("drops addColumn suggestions that would collide with an existing column name", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "addColumn", name: "Flag", formula: "1" },
          { type: "take", count: 5 },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Flag", type: "number" }], []);
    const ops = await suggestQueryNextSteps("add flag", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 5 }]);
  });

  it("drops addColumn suggestions with invalid formulas", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          // M language style expression (unsupported)
          { type: "addColumn", name: "Flag", formula: "if [Region] = 'East' then 1 else 0" },
          // Valid formula expression
          { type: "addColumn", name: "Flag2", formula: "[Region] == 'East' ? 1 : 0" },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("add flag", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "addColumn", name: "Flag2", formula: "[Region] == 'East' ? 1 : 0" }]);
  });

  it("drops addColumn formulas that reference '_' (value formulas)", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "addColumn", name: "Bad", formula: "_" },
          { type: "addColumn", name: "Ok", formula: "[Region]" },
        ]),
      },
    });

    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const ops = await suggestQueryNextSteps("add", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "addColumn", name: "Ok", formula: "[Region]" }]);
  });

  it("drops renameColumn suggestions that would collide with an existing column name", async () => {
    chatMock.mockResolvedValue({
      message: {
        role: "assistant",
        content: JSON.stringify([
          { type: "renameColumn", oldName: "Region", newName: "Sales" },
          { type: "take", count: 5 },
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
    const ops = await suggestQueryNextSteps("rename", { query: baseQuery(), preview });
    expect(ops).toEqual([{ type: "take", count: 5 }]);
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
