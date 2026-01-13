import { describe, expect, it } from "vitest";

import { extractSheetSchema } from "./schema.js";
import { pickBestRegionForQuery } from "./queryAware.js";

describe("queryAware region selection", () => {
  it("picks the Region/Revenue table for 'revenue by region'", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["Region", "Revenue", null, null, "Category", "Cost"],
        ["North", 100, null, null, "Labor", 50],
        ["South", 200, null, null, "Materials", 70],
      ],
      tables: [
        { name: "Revenue Data", range: "A1:B3" },
        { name: "Cost Data", range: "E1:F3" },
      ],
    };

    const schema = extractSheetSchema(sheet);
    const picked = pickBestRegionForQuery(schema, "revenue by region");

    expect(picked).toEqual({ type: "table", index: 0, range: "Sheet1!A1:B3" });
  });

  it("picks the cost table for 'cost'", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["Region", "Revenue", null, null, "Category", "Cost"],
        ["North", 100, null, null, "Labor", 50],
        ["South", 200, null, null, "Materials", 70],
      ],
      tables: [
        { name: "Revenue Data", range: "A1:B3" },
        { name: "Cost Data", range: "E1:F3" },
      ],
    };

    const schema = extractSheetSchema(sheet);
    const picked = pickBestRegionForQuery(schema, "cost");

    expect(picked).toEqual({ type: "table", index: 1, range: "Sheet1!E1:F3" });
  });
});

