import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop } from "./helpers";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

test.describe("Desktop extension spreadsheet API", () => {
  test("Sheet.getRange/setRange round-trips values", async ({ page }) => {
    await gotoDesktop(page);

    const extensionApiUrl = viteFsUrl(path.join(repoRoot, "packages/extension-api/index.mjs"));

    const result = await page.evaluate(
      async ({ extensionApiUrl }) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

        // Ensure the host is booted (DesktopExtensionHostManager lazily loads extensions).
        if (!manager.ready) {
          await manager.loadBuiltInExtensions();
        }

        const commandId = "rangeExt.roundTrip";
        const manifest = {
          name: "range-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Range round trip" }] },
          permissions: ["ui.commands", "cells.read", "cells.write"],
        };

        const code = `
          import * as formula from ${JSON.stringify(extensionApiUrl)};
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId,
            )}, async () => {
              const sheet = await formula.sheets.getActiveSheet();
              await sheet.setRange("A1:B2", [[1,2],[3,4]]);
              const range = await sheet.getRange("A1:B2");
              return range.values;
            }));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);
        const extensionId = `${manifest.publisher}.${manifest.name}`;

        try {
          await manager.host.loadExtension({
            extensionId,
            extensionPath: "memory://range-ext/",
            manifest,
            mainUrl,
          });

          return await manager.host.executeCommand(commandId);
        } finally {
          try {
            await manager.host.unloadExtension(extensionId);
          } catch {
            // ignore cleanup failures
          }
          URL.revokeObjectURL(mainUrl);
        }
      },
      { extensionApiUrl },
    );

    expect(result).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });
});

