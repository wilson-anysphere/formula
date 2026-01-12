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

async function setRestrictedRangeClassification(
  page: Page,
  params: { documentId: string; sheetId: string; range: { start: { row: number; col: number }; end: { row: number; col: number } } },
): Promise<void> {
  await page.addInitScript(
    ({ documentId, sheetId, range }) => {
      const key = `dlp:classifications:${documentId}`;
      const record = {
        selector: { scope: "range", documentId, sheetId, range },
        classification: { level: "Restricted", labels: [] },
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(key, JSON.stringify([record]));
    },
    params,
  );
}

async function assertClipboardSupportedOrSkip(page: Page): Promise<void> {
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);

  const clipboardSupport = await page.evaluate(async () => {
    if (!globalThis.isSecureContext) return { supported: false, reason: "not a secure context" };
    if (!navigator.clipboard?.readText || !navigator.clipboard?.writeText) {
      return { supported: false, reason: "navigator.clipboard.readText/writeText not available" };
    }

    try {
      const marker = `__formula_clipboard_probe__${Math.random().toString(16).slice(2)}`;
      await navigator.clipboard.writeText(marker);
      const echoed = await navigator.clipboard.readText();
      return { supported: echoed === marker, reason: echoed === marker ? null : `mismatch: ${echoed}` };
    } catch (err: any) {
      return { supported: false, reason: String(err?.message ?? err) };
    }
  });

  test.skip(!clipboardSupport.supported, `Clipboard APIs are blocked: ${clipboardSupport.reason ?? ""}`);
}

test.describe("Extension clipboard DLP (taint tracking)", () => {
  test("blocks clipboard.writeText when the extension read-taint intersects a Restricted range", async ({ page }) => {
    const extensionId = "formula-test.dlp-clipboard-block";
    const commandId = "dlpClipboard.blocked";

    await grantExtensionPermissions(page, extensionId, ["ui.commands", "cells.read", "clipboard"]);
    await setRestrictedRangeClassification(page, {
      documentId: "local-workbook",
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    });

    await gotoDesktop(page);
    await assertClipboardSupportedOrSkip(page);

    // Move off the Restricted cell so selection-based DLP enforcement doesn't block this test.
    // This ensures the block is coming from the extension host's taint tracking + clipboardWriteGuard.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.activateCell({ sheetId, row: 0, col: 1 }); // B1
    });

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;
    await page.evaluate(async (marker) => {
      await navigator.clipboard.writeText(marker);
    }, marker);

    const result = await page.evaluate(
      async ({ extensionId, commandId }) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager?.host) throw new Error("Missing window.__formulaExtensionHostManager.host");

        // Ensure the sheet has some content; DLP enforcement is based on classification metadata,
        // but having a real value makes the scenario more realistic.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, "Secret");

        const manifest = {
          name: "dlp-clipboard-block",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "DLP blocked clipboard write" }] },
          permissions: ["ui.commands", "cells.read", "clipboard"],
        };

        const code = `
          export async function activate(context) {
            const api = globalThis[Symbol.for("formula.extensionApi.api")];
            if (!api) throw new Error("Missing Formula extension API runtime");
            context.subscriptions.push(await api.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              // Taint the A1 cell, then attempt to write to clipboard.
              await api.cells.getCell(0, 0);
              await api.clipboard.writeText("leak");
              return "wrote";
            }));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);

        try {
          await manager.host.loadExtension({
            extensionId,
            extensionPath: "memory://dlp-clipboard-block/",
            manifest,
            mainUrl,
          });

          let errorMessage = "";
          try {
            await manager.host.executeCommand(commandId);
          } catch (err: any) {
            errorMessage = String(err?.message ?? err);
          }

          return { errorMessage };
        } finally {
          try {
            await manager.host.unloadExtension(extensionId);
          } catch {
            // ignore cleanup failures
          }
          URL.revokeObjectURL(mainUrl);
        }
      },
      { extensionId, commandId },
    );

    expect(result.errorMessage).toContain("Clipboard copy is blocked");

    await expect(page.locator('[data-testid="toast"][data-type="error"]')).toContainText("Clipboard copy is blocked");

    const clipboardText = await page.evaluate(async () => await navigator.clipboard.readText());
    expect(clipboardText).toBe(marker);
  });

  test("allows clipboard.writeText when the extension did not read any cells (even if the selection is Restricted)", async ({
    page,
  }) => {
    const extensionId = "formula-test.dlp-clipboard-allow";
    const commandId = "dlpClipboard.allowed";

    await grantExtensionPermissions(page, extensionId, ["ui.commands", "clipboard"]);
    await setRestrictedRangeClassification(page, {
      documentId: "local-workbook",
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    });

    await gotoDesktop(page);
    await assertClipboardSupportedOrSkip(page);

    // Move the selection away and back before loading the extension. This ensures any host-side
    // bookkeeping that keys off "selection changed at least once" is exercised, while still
    // keeping the extension untainted (it isn't loaded yet).
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.activateCell({ sheetId, row: 0, col: 1 }); // B1
      app.activateCell({ sheetId, row: 0, col: 0 }); // A1 (Restricted)
    });

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;
    await page.evaluate(async (marker) => {
      await navigator.clipboard.writeText(marker);
    }, marker);

    const result = await page.evaluate(
      async ({ extensionId, commandId }) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const manager: any = (window as any).__formulaExtensionHostManager;
        if (!manager?.host) throw new Error("Missing window.__formulaExtensionHostManager.host");

        const manifest = {
          name: "dlp-clipboard-allow",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "DLP allowed clipboard write" }] },
          permissions: ["ui.commands", "clipboard"],
        };

        const code = `
          export async function activate(context) {
            const api = globalThis[Symbol.for("formula.extensionApi.api")];
            if (!api) throw new Error("Missing Formula extension API runtime");
            context.subscriptions.push(await api.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              // No cell reads => no taint => should not trigger clipboard DLP enforcement.
              await api.clipboard.writeText("hello-from-extension");
              return "wrote";
            }));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);

        try {
          await manager.host.loadExtension({
            extensionId,
            extensionPath: "memory://dlp-clipboard-allow/",
            manifest,
            mainUrl,
          });

          const out = await manager.host.executeCommand(commandId);
          return { out };
        } finally {
          try {
            await manager.host.unloadExtension(extensionId);
          } catch {
            // ignore cleanup failures
          }
          URL.revokeObjectURL(mainUrl);
        }
      },
      { extensionId, commandId },
    );

    expect(result.out).toBe("wrote");

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe("hello-from-extension");
  });

  test("sampleHello.copySumToClipboard is blocked when the selection is Restricted", async ({ page }) => {
    const extensionId = "formula.sample-hello";
    await grantExtensionPermissions(page, extensionId, ["ui.commands", "cells.read", "clipboard"]);
    await setRestrictedRangeClassification(page, {
      documentId: "local-workbook",
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
    });

    await gotoDesktop(page);
    await assertClipboardSupportedOrSkip(page);

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;
    await page.evaluate(async (marker) => {
      await navigator.clipboard.writeText(marker);
    }, marker);

    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const manager: any = (window as any).__formulaExtensionHostManager;
      if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

      // Ensure the host is booted (DesktopExtensionHostManager lazily loads extensions).
      if (!manager.ready) {
        await manager.loadBuiltInExtensions();
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();

      app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, 1);
      app.getDocument().setCellValue(sheetId, { row: 0, col: 1 }, 2);
      app.getDocument().setCellValue(sheetId, { row: 1, col: 0 }, 3);
      app.getDocument().setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({ sheetId, range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } });

      let errorMessage = "";
      try {
        await manager.host.executeCommand("sampleHello.copySumToClipboard");
      } catch (err: any) {
        errorMessage = String(err?.message ?? err);
      }

      return { errorMessage };
    });

    expect(result.errorMessage).toContain("Clipboard copy is blocked");

    const clipboardText = await page.evaluate(async () => await navigator.clipboard.readText());
    expect(clipboardText).toBe(marker);
  });
});
