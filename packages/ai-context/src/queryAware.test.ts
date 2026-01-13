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

  it("returns null when no region matches the query", () => {
    const schema = {
      name: "Sheet1",
      tables: [
        {
          name: "Unrelated",
          range: "Sheet1!A1:B2",
          columns: [{ name: "Foo", type: "string", sampleValues: [] }],
          rowCount: 1,
        },
      ],
      namedRanges: [],
      dataRegions: [
        {
          range: "Sheet1!A1:B2",
          hasHeader: true,
          headers: ["Foo"],
          inferredColumnTypes: ["string"],
          rowCount: 1,
          columnCount: 1,
        },
      ],
    } as any;

    expect(pickBestRegionForQuery(schema, "revenue")).toBeNull();
  });

  it("prefers tables over dataRegions when scores tie", () => {
    const schema = {
      name: "Sheet1",
      tables: [
        {
          name: "Data",
          range: "Sheet1!A1:B3",
          columns: [
            { name: "Region", type: "string", sampleValues: [] },
            { name: "Revenue", type: "number", sampleValues: [] },
          ],
          rowCount: 2,
        },
      ],
      namedRanges: [],
      dataRegions: [
        {
          range: "Sheet1!D1:E3",
          hasHeader: true,
          headers: ["Region", "Revenue"],
          inferredColumnTypes: ["string", "number"],
          rowCount: 2,
          columnCount: 2,
        },
      ],
    } as any;

    const picked = pickBestRegionForQuery(schema, "revenue by region");
    expect(picked).toEqual({ type: "table", index: 0, range: "Sheet1!A1:B3" });
  });
});
