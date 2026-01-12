import { describe, expect, it } from "vitest";

import { EXCEL_MAX_SHEET_NAME_LEN, getSheetNameValidationErrorMessage } from "../sheetNameValidation";

describe("sheetNameValidation", () => {
  it("rejects blank/whitespace-only names", () => {
    expect(getSheetNameValidationErrorMessage("")).toBe("sheet name cannot be blank");
    expect(getSheetNameValidationErrorMessage("   ")).toBe("sheet name cannot be blank");
    expect(getSheetNameValidationErrorMessage("\n\t")).toBe("sheet name cannot be blank");
  });

  it("rejects leading/trailing apostrophe", () => {
    expect(getSheetNameValidationErrorMessage("'Budget")).toBe("sheet name cannot begin or end with an apostrophe");
    expect(getSheetNameValidationErrorMessage("Budget'")).toBe("sheet name cannot begin or end with an apostrophe");
  });

  it("rejects invalid characters", () => {
    expect(getSheetNameValidationErrorMessage("Bad:Name")).toBe("sheet name contains invalid character `:`");
    expect(getSheetNameValidationErrorMessage("Bad/Name")).toBe("sheet name contains invalid character `/`");
    expect(getSheetNameValidationErrorMessage("Bad[Name")).toBe("sheet name contains invalid character `[`");
  });

  it("enforces the 31 character limit in UTF-16 code units (JS string.length)", () => {
    // 🙂 is outside the BMP and takes two UTF-16 code units in JS (`"🙂".length === 2`).
    const maxOk = `${"a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 2)}🙂`;
    expect(maxOk.length).toBe(EXCEL_MAX_SHEET_NAME_LEN);
    expect(getSheetNameValidationErrorMessage(maxOk)).toBe(null);

    const tooLong = `${"a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 1)}🙂`;
    expect(tooLong.length).toBe(EXCEL_MAX_SHEET_NAME_LEN + 1);
    expect(getSheetNameValidationErrorMessage(tooLong)).toBe(`sheet name cannot exceed ${EXCEL_MAX_SHEET_NAME_LEN} characters`);
  });

  it("enforces workbook-wide uniqueness case-insensitively (Unicode NFKC + uppercasing)", () => {
    expect(getSheetNameValidationErrorMessage("budget", { existingNames: ["Budget"] })).toBe("sheet name already exists");
  });
});

