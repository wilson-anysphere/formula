import { describe, expect, it } from "vitest";

import { formatA1Cell, formatA1Range, parseA1Cell } from "./a1";

describe("ai-tools A1 sheet-name formatting", () => {
  it("formats unquoted identifier-like sheet names without quotes", () => {
    expect(formatA1Cell({ sheet: "Sheet1", row: 1, col: 1 })).toBe("Sheet1!A1");
    expect(formatA1Cell({ sheet: "Sheet.Name", row: 1, col: 1 })).toBe("Sheet.Name!A1");
  });

  it("quotes sheet names containing spaces/special characters", () => {
    expect(formatA1Cell({ sheet: "My Sheet", row: 1, col: 1 })).toBe("'My Sheet'!A1");
  });

  it("quotes reserved/ambiguous sheet names (TRUE, A1, R1C1, leading digits)", () => {
    expect(formatA1Cell({ sheet: "TRUE", row: 1, col: 1 })).toBe("'TRUE'!A1");
    expect(formatA1Cell({ sheet: "A1", row: 1, col: 1 })).toBe("'A1'!A1");
    expect(formatA1Cell({ sheet: "R1C1", row: 1, col: 1 })).toBe("'R1C1'!A1");
    expect(formatA1Cell({ sheet: "1Sheet", row: 1, col: 1 })).toBe("'1Sheet'!A1");
  });

  it("escapes apostrophes when quoting", () => {
    expect(formatA1Cell({ sheet: "O'Brien", row: 1, col: 1 })).toBe("'O''Brien'!A1");
  });

  it("parses quoted sheet prefixes", () => {
    expect(parseA1Cell("'My Sheet'!B2")).toEqual({ sheet: "My Sheet", row: 2, col: 2 });
  });

  it("formats ranges with sheet prefixes", () => {
    expect(
      formatA1Range({
        sheet: "TRUE",
        startRow: 1,
        startCol: 1,
        endRow: 2,
        endCol: 2,
      }),
    ).toBe("'TRUE'!A1:B2");
  });
});

