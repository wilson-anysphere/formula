import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("Organize Sheets ribbon wiring", () => {
  it("routes the ribbon command id through CommandRegistry to openOrganizeSheets()", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const commandsPath = fileURLToPath(new URL("../../commands/registerDesktopCommands.ts", import.meta.url));
    const main = stripComments(readFileSync(mainTsPath, "utf8"));
    const commands = stripComments(readFileSync(commandsPath, "utf8"));

    // The ribbon schema uses `home.cells.format.organizeSheets`. Ensure it's registered as a
    // real CommandRegistry command (no main.ts switch-case wiring) and delegates to the
    // `sheetStructureHandlers` hook.
    const registerMatch = commands.match(
      /\bregisterBuiltinCommand\s*\(\s*["']home\.cells\.format\.organizeSheets["']/,
    );
    expect(registerMatch).not.toBeNull();
    const registerIndex = registerMatch?.index ?? -1;
    expect(registerIndex).toBeGreaterThanOrEqual(0);
    // Ensure the handler reference is within the same registration block (i.e. before the next command registration),
    // but avoid brittle fixed-length windows.
    const nextRegisterIndex = commands.indexOf("registerBuiltinCommand", registerIndex + 1);
    const blockEnd = nextRegisterIndex >= 0 ? nextRegisterIndex : commands.length;
    const registrationBlock = commands.slice(registerIndex, blockEnd);
    expect(registrationBlock).toMatch(/\bsheetStructureHandlers\s*\?\.\s*openOrganizeSheets\b/);

    // Ensure `main.ts` passes the handler into registerDesktopCommands (so the command can open the dialog).
    expect(main).toMatch(/\bsheetStructureHandlers\s*:\s*\{[\s\S]{0,800}?\bopenOrganizeSheets\b/);

    // Ensure the helper exists and delegates to `openOrganizeSheetsDialog`.
    const fnMatch = main.match(/(?:function\s+openOrganizeSheets\s*\(|const\s+openOrganizeSheets\s*=\s*\(\)\s*=>)/);
    expect(fnMatch).not.toBeNull();
    const fnIndex = fnMatch?.index ?? -1;
    expect(fnIndex).toBeGreaterThanOrEqual(0);
    expect(main.slice(fnIndex, fnIndex + 1600)).toContain("openOrganizeSheetsDialog(");
  });
});
