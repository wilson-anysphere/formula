import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

describe("Organize Sheets ribbon wiring", () => {
  it("routes the ribbon command id to openOrganizeSheets()", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const source = readFileSync(mainTsPath, "utf8");

    // Ensure the ribbon command id is explicitly handled and opens the dialog.
    const caseNeedle = 'case "home.cells.format.organizeSheets":';
    const caseIndex = source.indexOf(caseNeedle);
    expect(caseIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(caseIndex, caseIndex + 250)).toContain("openOrganizeSheets()");

    // Ensure the helper exists and delegates to `openOrganizeSheetsDialog`.
    const fnNeedle = "function openOrganizeSheets(): void {";
    const fnIndex = source.indexOf(fnNeedle);
    expect(fnIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(fnIndex, fnIndex + 1200)).toContain("openOrganizeSheetsDialog(");
  });
});
