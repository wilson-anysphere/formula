import { describe, expect, it, vi } from "vitest";

// `ContextManager.buildContext()` needs a sheet schema for prompt context and also indexes the
// sheet for RAG retrieval. Historically both steps extracted the schema separately. This test
// ensures schema extraction happens only once per `buildContext()` call.
vi.mock("./schema.js", async () => {
  const actual = await vi.importActual<typeof import("./schema.js")>("./schema.js");
  return {
    ...actual,
    extractSheetSchema: vi.fn(actual.extractSheetSchema),
  };
});

import { ContextManager } from "./contextManager.js";
import { extractSheetSchema } from "./schema.js";

describe("ContextManager.buildContext schema extraction", () => {
  it("extracts the sheet schema only once per buildContext call", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      redactor: (text: string) => text,
    });

    const sheet = {
      name: "Sheet1",
      values: [
        ["Region", "Revenue"],
        ["North", 1000],
        ["South", 2000],
      ],
    };

    vi.mocked(extractSheetSchema).mockClear();
    await cm.buildContext({ sheet, query: "revenue" });

    expect(vi.mocked(extractSheetSchema)).toHaveBeenCalledTimes(1);
  });
});

