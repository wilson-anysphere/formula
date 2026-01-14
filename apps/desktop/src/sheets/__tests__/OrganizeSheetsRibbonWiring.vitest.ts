import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

describe("Organize Sheets ribbon wiring", () => {
  it("routes the ribbon command id to openOrganizeSheets()", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const source = readFileSync(mainTsPath, "utf8");

    // Ensure the ribbon command id is explicitly handled and opens the dialog.
    // Be tolerant of minor formatting differences (single vs double quotes, whitespace).
    const caseMatch = source.match(/case\s+["']home\.cells\.format\.organizeSheets["']\s*:/);
    expect(caseMatch).not.toBeNull();
    const caseIndex = caseMatch?.index ?? -1;
    expect(caseIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(caseIndex, caseIndex + 300)).toMatch(/openOrganizeSheets\s*\(/);

    // Ensure the helper exists and delegates to `openOrganizeSheetsDialog`.
    const fnMatch = source.match(/(?:function\s+openOrganizeSheets\s*\(|const\s+openOrganizeSheets\s*=\s*\(\)\s*=>)/);
    expect(fnMatch).not.toBeNull();
    const fnIndex = fnMatch?.index ?? -1;
    expect(fnIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(fnIndex, fnIndex + 1600)).toContain("openOrganizeSheetsDialog(");
  });
});
