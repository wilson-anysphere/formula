import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function grantExtensionPermissions(page: Page, extensionId: string, permissions: string[]): Promise<void> {
  await page.addInitScript(
    ({ extensionId, permissions }) => {
      const key = "formula.extensionHost.permissions";
      const existing = (() => {
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : {};
        } catch {
          return {};
        }
      })();

      existing[extensionId] = {
        ...(existing[extensionId] ?? {}),
        ...Object.fromEntries(permissions.map((perm) => [perm, true])),
      };

      localStorage.setItem(key, JSON.stringify(existing));
    },
    { extensionId, permissions },
  );
}

test.describe("Desktop extension spreadsheet API", () => {
  test("Sheet.getRange/setRange round-trips values", async ({ page }) => {
    await grantExtensionPermissions(page, "formula-test.range-ext", ["ui.commands", "cells.read", "cells.write"]);
    await gotoDesktop(page);

    const result = await page.evaluate(
      async () => {
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
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!formula) throw new Error("Missing formula extension API runtime");
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
    );

    expect(result).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  test("formula.sheets.createSheet creates a sheet and Sheet.getRange/setRange work on it", async ({ page }) => {
    await grantExtensionPermissions(page, "formula-test.sheet-create-ext", [
      "ui.commands",
      "sheets.manage",
      "cells.read",
      "cells.write",
    ]);
    await gotoDesktop(page);

    const result = await page.evaluate(
      async () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

        if (!manager.ready) {
          await manager.loadBuiltInExtensions();
        }

        const commandId = "sheetExt.createAndRoundTrip";
        const manifest = {
          name: "sheet-create-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Create sheet + range round trip" }] },
          permissions: ["ui.commands", "sheets.manage", "cells.read", "cells.write"],
        };

        const code = `
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!formula) throw new Error("Missing formula extension API runtime");
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId,
            )}, async () => {
              const sheet = await formula.sheets.createSheet("Data");
              await sheet.setRange("A1:B2", [[1,2],[3,4]]);
              const range = await sheet.getRange("A1:B2");
              const active = await formula.sheets.getActiveSheet();
              return { values: range.values, sheetName: sheet.name, activeName: active.name };
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
            extensionPath: "memory://sheet-create-ext/",
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
    );

    expect(result.sheetName).toBe("Data");
    expect(result.activeName).toBe("Data");
    expect(result.values).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  test("formula.sheets.activateSheet activates the real UI sheet", async ({ page }) => {
    await grantExtensionPermissions(page, "formula-test.sheet-ext", ["ui.commands", "sheets.manage"]);
    await gotoDesktop(page);

    // Create a second sheet in the underlying DocumentController model so we can activate it.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      doc.setCellValue("Sheet2", { row: 0, col: 0 }, "Hello");
    });

    const result = await page.evaluate(
      async () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

        if (!manager.ready) {
          await manager.loadBuiltInExtensions();
        }

        const commandId = "sheetExt.activate";
        const manifest = {
          name: "sheet-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Activate Sheet2" }] },
          permissions: ["ui.commands", "sheets.manage"],
        };

        const code = `
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!formula) throw new Error("Missing formula extension API runtime");
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId,
            )}, async () => {
              const sheet = await formula.sheets.activateSheet("Sheet2");
              const active = await formula.sheets.getActiveSheet();
              return { activated: { id: sheet.id, name: sheet.name }, active: { id: active.id, name: active.name } };
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
            extensionPath: "memory://sheet-ext/",
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
    );

    expect(result.active.name).toBe("Sheet2");

    const activeSheetId = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      return app.getCurrentSheetId();
    });

    expect(activeSheetId).toBe("Sheet2");
  });

  test("formula.events.onSheetActivated fires when an extension activates a sheet", async ({ page }) => {
    await grantExtensionPermissions(page, "formula-test.sheet-events-ext", ["ui.commands", "sheets.manage"]);
    await gotoDesktop(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      doc.setCellValue("Sheet2", { row: 0, col: 0 }, "Hello");
    });

    const result = await page.evaluate(
      async () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

        if (!manager.ready) {
          await manager.loadBuiltInExtensions();
        }

        const commandId = "sheetExt.onActivated";
        const manifest = {
          name: "sheet-events-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Wait for sheetActivated" }] },
          permissions: ["ui.commands", "sheets.manage"],
        };

        const code = `
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!formula) throw new Error("Missing formula extension API runtime");
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId,
            )}, async () => {
              const activated = new Promise((resolve) => {
                const disp = formula.events.onSheetActivated((e) => {
                  try { disp.dispose(); } catch {}
                  resolve(e?.sheet?.name ?? null);
                });
                // Safety timeout so the command never hangs the e2e run.
                setTimeout(() => {
                  try { disp.dispose(); } catch {}
                  resolve(null);
                }, 2000);
              });

              await formula.sheets.activateSheet("Sheet2");
              const active = await formula.sheets.getActiveSheet();
              const fired = await activated;
              return { fired, active: { id: active.id, name: active.name } };
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
            extensionPath: "memory://sheet-events-ext/",
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
    );

    expect(result.active.name).toBe("Sheet2");
    expect(result.fired).toBe("Sheet2");
  });
});
