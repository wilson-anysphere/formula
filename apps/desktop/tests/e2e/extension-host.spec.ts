import { expect, test } from "@playwright/test";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop, openExtensionsPanel } from "./helpers";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

test.describe("BrowserExtensionHost", () => {
  test("loads sample extension in a Worker and can run sumSelection", async ({ page }) => {
    await gotoDesktop(page);

    const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));

    const result = await page.evaluate(
      async ({ manifestUrl, hostModuleUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

        doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
        doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
        doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
        doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

        app.selectRange({
          sheetId,
          range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }
        });

        function normalizeCellValue(value: unknown) {
          if (typeof value === "string") return value;
          if (typeof value === "number") return value;
          if (typeof value === "boolean") return value;
          return null;
        }

        const spreadsheetApi = {
          async getActiveSheet() {
            return { id: sheetId, name: sheetId };
          },
          async getSelection() {
            const range = app.getSelectionRanges()[0];
            const values = [];
            for (let r = range.startRow; r <= range.endRow; r++) {
              const cols = [];
              for (let c = range.startCol; c <= range.endCol; c++) {
                const cell = doc.getCell(sheetId, { row: r, col: c });
                cols.push(normalizeCellValue(cell.value));
              }
              values.push(cols);
            }
            return { ...range, values };
          },
          async getCell(row: number, col: number) {
            const cell = doc.getCell(sheetId, { row, col });
            return normalizeCellValue(cell.value);
          },
          async setCell(row: number, col: number, value: unknown) {
            doc.setCellValue(sheetId, { row, col }, value);
          }
        };

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi,
          permissionPrompt: async () => true,
          // Vite rewrites `@formula/extension-api` into `/@fs/...` URLs, which fail the
          // strict import preflight. The import sandbox is exercised in unit tests; for
          // this e2e suite we disable preflight so we can validate host behavior.
          sandbox: { strictImports: false },
        });

        await host.loadExtensionFromUrl(manifestUrl);

        const sum = await host.executeCommand("sampleHello.sumSelection");
        const a3 = await app.getCellValueA1("A3");
        await host.dispose();

        return { sum, a3 };
      },
      { manifestUrl, hostModuleUrl }
    );

    expect(result.sum).toBe(10);
    expect(result.a3).toBe("10");
  });

  test("activation context includes storage paths and matches formula.context", async ({ page }) => {
    await gotoDesktop(page);

    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
    const extensionApiUrl = viteFsUrl(path.join(repoRoot, "packages/extension-api/index.mjs"));

    const result = await page.evaluate(
      async ({ hostModuleUrl, extensionApiUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        // The extension entrypoint is loaded via `blob:` URL, so its module resolution base is
        // non-hierarchical. Convert Vite's `/@fs/...` path into an absolute http(s) URL so
        // `import` inside the blob-backed worker can resolve it.
        const extensionApiAbsoluteUrl = new URL(extensionApiUrl, location.href).href;

        const commandId = "ctxExt.read";
        const manifest = {
          name: "ctx-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Read Context" }] },
          permissions: ["ui.commands"],
        };

        const code = `
          import * as formula from ${JSON.stringify(extensionApiAbsoluteUrl)};
          export async function activate(context) {
            const snapshot = {
              ctx: {
                extensionId: context.extensionId,
                extensionPath: context.extensionPath,
                extensionUri: context.extensionUri,
                globalStoragePath: context.globalStoragePath,
                workspaceStoragePath: context.workspaceStoragePath
              },
              api: {
                extensionId: formula.context.extensionId,
                extensionPath: formula.context.extensionPath,
                extensionUri: formula.context.extensionUri,
                globalStoragePath: formula.context.globalStoragePath,
                workspaceStoragePath: formula.context.workspaceStoragePath
              }
            };

            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId
            )}, async () => snapshot));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi: {},
          permissionPrompt: async () => true,
          sandbox: { strictImports: false },
        });

        await host.loadExtension({
          extensionId: `${manifest.publisher}.${manifest.name}`,
          extensionPath: "memory://ctx-ext/",
          manifest,
          mainUrl,
        });

        const snapshot = await host.executeCommand(commandId);
        await host.dispose();
        URL.revokeObjectURL(mainUrl);

        return snapshot;
      },
      { hostModuleUrl, extensionApiUrl }
    );

    expect(result.ctx.extensionId).toBe("formula-test.ctx-ext");
    expect(result.ctx.extensionPath).toBe("memory://ctx-ext/");
    expect(result.ctx.extensionUri).toBe("memory://ctx-ext/");
    expect(String(result.ctx.globalStoragePath)).toContain("globalStorage");
    expect(String(result.ctx.workspaceStoragePath)).toContain("workspaceStorage");

    // Extension API should reflect the same values as the activation context.
    expect(result.api).toEqual(result.ctx);
  });

  test("network.fetch is permission gated in the browser host", async ({ page }) => {
    const server = http.createServer((req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello");
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : null;
    if (!port) throw new Error("Failed to allocate test port");

    try {
      await gotoDesktop(page);

      const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
      const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
      const url = `http://127.0.0.1:${port}/`;

      const result = await page.evaluate(
        async ({ manifestUrl, hostModuleUrl, url }) => {
          const { BrowserExtensionHost } = await import(hostModuleUrl);

          const host = new BrowserExtensionHost({
            engineVersion: "1.0.0",
            spreadsheetApi: {
              async getSelection() {
                return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
              },
              async getCell() {
                return null;
              },
              async setCell() {
                // noop
              },
            },
            permissionPrompt: async () => true,
            sandbox: { strictImports: false },
          });

          await host.loadExtensionFromUrl(manifestUrl);
          const text = await host.executeCommand("sampleHello.fetchText", url);
          const messages = host.getMessages();
          await host.dispose();
          return { text, messages };
        },
        { manifestUrl, hostModuleUrl, url }
      );

    expect(result.text).toBe("hello");
    expect(result.messages.some((m: any) => String(m.message).includes("Fetched: hello"))).toBe(true);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
  });

  test("delegates workbook.createWorkbook to the provided spreadsheetApi and returns a workbook snapshot", async ({
    page,
  }) => {
    await gotoDesktop(page);

    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
    const extensionApiUrl = viteFsUrl(path.join(repoRoot, "packages/extension-api/index.mjs"));

    const result = await page.evaluate(
      async ({ hostModuleUrl, extensionApiUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        // The extension entrypoint is loaded via `blob:` URL, so its module resolution base is
        // non-hierarchical. Convert Vite's `/@fs/...` path into an absolute http(s) URL so
        // `import` inside the blob-backed worker can resolve it.
        const extensionApiAbsoluteUrl = new URL(extensionApiUrl, location.href).href;

        const commandId = "wbExt.createAndRead";
        const manifest = {
          name: "wb-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "Create Workbook + Read" }] },
          permissions: ["ui.commands", "workbook.manage"],
        };

        const code = `
          import * as formula from ${JSON.stringify(extensionApiAbsoluteUrl)};
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId,
            )}, async () => {
              await formula.workbook.createWorkbook();
              const wb = await formula.workbook.getActiveWorkbook();
              return { name: wb.name, path: wb.path, sheets: wb.sheets, activeSheet: wb.activeSheet };
            }));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);

        const sheet = { id: "Sheet1", name: "Sheet1" };
        const state = { name: "InitialWorkbook", path: null };

        const spreadsheetApi = {
          listSheets() {
            return [sheet];
          },
          getActiveSheet() {
            return sheet;
          },
          async getActiveWorkbook() {
            return { name: state.name, path: state.path };
          },
          async createWorkbook() {
            state.name = "CreatedWorkbook";
            state.path = null;
          },
          async getSelection() {
            return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
          },
          async getCell() {
            return null;
          },
          async setCell() {
            // noop
          },
        };

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi,
          permissionPrompt: async () => true,
          sandbox: { strictImports: false },
        });

        await host.loadExtension({
          extensionId: `${manifest.publisher}.${manifest.name}`,
          extensionPath: "memory://wb-ext/",
          manifest,
          mainUrl,
        });

        const workbook = await host.executeCommand(commandId);
        await host.dispose();
        URL.revokeObjectURL(mainUrl);

        return workbook;
      },
      { hostModuleUrl, extensionApiUrl },
    );

    expect(result.name).toBe("CreatedWorkbook");
    expect(result.path).toBeNull();
    expect(Array.isArray(result.sheets)).toBe(true);
    expect(result.sheets.length).toBeGreaterThan(0);
    expect(result.sheets[0]).toEqual({ id: "Sheet1", name: "Sheet1" });
    expect(result.activeSheet).toEqual({ id: "Sheet1", name: "Sheet1" });
  });

  test("clipboard.writeText writes to the system clipboard (desktop adapter)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

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

    const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
    const extensionHostManagerUrl = viteFsUrl(
      path.join(repoRoot, "apps/desktop/src/extensions/extensionHostManager.ts"),
    );

    const result = await page.evaluate(
      async ({ manifestUrl, extensionHostManagerUrl }) => {
        const { DesktopExtensionHostManager } = await import(extensionHostManagerUrl);

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

        doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
        doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
        doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
        doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

        app.selectRange({
          sheetId,
          range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        });

        function normalizeCellValue(value: unknown) {
          if (typeof value === "string") return value;
          if (typeof value === "number") return value;
          if (typeof value === "boolean") return value;
          return null;
        }

        const spreadsheetApi = {
          async getActiveSheet() {
            return { id: sheetId, name: sheetId };
          },
          async getSelection() {
            const range = app.getSelectionRanges()[0];
            const values = [];
            for (let r = range.startRow; r <= range.endRow; r++) {
              const cols = [];
              for (let c = range.startCol; c <= range.endCol; c++) {
                const cell = doc.getCell(sheetId, { row: r, col: c });
                cols.push(normalizeCellValue(cell.value));
              }
              values.push(cols);
            }
            return { ...range, values };
          },
          async getCell(row: number, col: number) {
            const cell = doc.getCell(sheetId, { row, col });
            return normalizeCellValue(cell.value);
          },
          async setCell(row: number, col: number, value: unknown) {
            doc.setCellValue(sheetId, { row, col }, value);
          },
        };

        const manager = new DesktopExtensionHostManager({
          engineVersion: "1.0.0",
          spreadsheetApi,
          clipboardApi: {
            readText: async () => navigator.clipboard.readText(),
            writeText: async (next: string) => navigator.clipboard.writeText(String(next ?? "")),
          },
          uiApi: {},
          permissionPrompt: async () => true,
        });

        await manager.host.loadExtensionFromUrl(manifestUrl);

        const sum = await manager.host.executeCommand("sampleHello.copySumToClipboard");
        const clipboardText = await navigator.clipboard.readText();
        await manager.host.dispose();

        return { sum, clipboardText };
      },
      { manifestUrl, extensionHostManagerUrl },
    );

    expect(result.sum).toBe(10);
    expect(result.clipboardText).toBe("10");
  });

  test("clipboard.writeText is blocked by selection DLP (desktop runtime adapter)", async ({
    page,
  }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

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

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.beginBatch({ label: "Seed selection DLP clipboard test" });
      doc.setCellValue(sheetId, "A1", "RESTRICTED");
      doc.setCellValue(sheetId, "B1", "SAFE");
      doc.endBatch();
      app.refresh();

      const docIdParam = new URL(window.location.href).searchParams.get("docId");
      const docId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam.trim() : null;
      const workbookId = docId ?? "local-workbook";

      // Mark A1 as Restricted.
      const record = {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 0, col: 0 },
        classification: { level: "Restricted", labels: [] },
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(`dlp:classifications:${workbookId}`, JSON.stringify([record]));

      // Pre-grant permissions for a small in-memory test extension.
      const extensionId = "e2e.dlp-selection";
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
        "ui.commands": true,
        clipboard: true,
      };
      localStorage.setItem(key, JSON.stringify(existing));
    });
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;
    await page.evaluate(async (marker) => await navigator.clipboard.writeText(marker), marker);

    // Move selection away and back before loading the extension. This ensures any host-side
    // bookkeeping that keys off "selection changed at least once" is exercised, and leaves the
    // UI selection on the Restricted cell so selection-based DLP blocks the clipboard write.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.activateCell({ sheetId, row: 0, col: 1 }); // B1
      app.activateCell({ sheetId, row: 0, col: 0 }); // A1 (Restricted)
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Clear pre-existing toasts so we can assert that DLP doesn't fire.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const host: any = (window as any).__formulaExtensionHost;
      if (!host) throw new Error("Missing window.__formulaExtensionHost (desktop e2e harness)");

      const extensionId = "e2e.dlp-selection";
      const commandId = "dlpSelection.writeConstant";
      const mainSource = `
        const api = globalThis[Symbol.for("formula.extensionApi.api")];
        export async function activate(context) {
          context.subscriptions.push(
            await api.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              await api.clipboard.writeText("EVIL");
              return "ok";
            }),
          );
        }
      `;

      const mainUrl = URL.createObjectURL(new Blob([mainSource], { type: "text/javascript" }));
      const manifest = {
        name: "dlp-selection",
        publisher: "e2e",
        version: "1.0.0",
        engines: { formula: "^1.0.0" },
        main: "./dist/extension.mjs",
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Write clipboard constant", category: "DLP Test" }] },
        permissions: ["clipboard", "ui.commands"],
      };

      try {
        await host.unloadExtension(extensionId);
      } catch {
        // ignore
      }

      await host.loadExtension({ extensionId, extensionPath: "memory://e2e/dlp-selection", manifest, mainUrl });

      try {
        await host.executeCommand(commandId);
        return { ok: true };
      } catch (err) {
        return { ok: false, message: String((err as any)?.message ?? err) };
      } finally {
        URL.revokeObjectURL(mainUrl);
        try {
          await host.unloadExtension(extensionId);
        } catch {
          // ignore
        }
      }
    });

    expect(result.ok).toBe(false);
    expect(result.message).toContain("Clipboard copy is blocked");
    await expect(page.getByTestId("toast-root")).toContainText("Clipboard copy is blocked");

    // Ensure the clipboard was not modified.
    await expect.poll(() => page.evaluate(async () => await navigator.clipboard.readText())).toBe(marker);
  });

  test("clipboard.writeText is blocked when Restricted data is received via events.onSelectionChanged", async ({
    page,
  }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);
    await openExtensionsPanel(page);

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

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;

    const result = await page.evaluate(
      async ({ marker }) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const hostManager: any = (window as any).__formulaExtensionHostManager;
        if (!hostManager) throw new Error("Missing window.__formulaExtensionHostManager (desktop runtime)");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const host: any = (window as any).__formulaExtensionHost;
        if (!host) throw new Error("Missing window.__formulaExtensionHost (desktop runtime)");

        // Ensure extension permission prompts are non-interactive in e2e.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).__formulaPermissionPrompt = async () => true;

        // Clear pre-existing toasts so the test can assert on the new DLP error.
        document.getElementById("toast-root")?.replaceChildren();

        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

        doc.setCellValue(sheetId, { row: 0, col: 0 }, "Secret"); // A1
        doc.setCellValue(sheetId, { row: 0, col: 1 }, "Public"); // B1

        const docIdParam = new URL(window.location.href).searchParams.get("docId");
        const docId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam.trim() : null;
        const workbookId = docId ?? "local-workbook";

        // Mark A1 as Restricted so DLP blocks clipboard.copy and extension clipboard.writeText.
        try {
          const key = `dlp:classifications:${workbookId}`;
          localStorage.setItem(
            key,
            JSON.stringify([
              {
                selector: {
                  scope: "range",
                  documentId: workbookId,
                  sheetId,
                  range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
                },
                classification: { level: "Restricted", labels: [] },
                updatedAt: new Date().toISOString(),
              },
            ]),
          );
        } catch {
          // ignore localStorage failures; test will fail if DLP isn't applied.
        }

        await navigator.clipboard.writeText(marker);

        // Ensure the desktop extension host is fully started (built-ins loaded + startup hooks run).
        await hostManager.loadBuiltInExtensions();

        // Load an extension that only reads cell values via `events.onSelectionChanged` and then
        // attempts to write those values to the system clipboard.
        const extensionId = `formula-test.event-taint-${Math.random().toString(16).slice(2)}`;
        const commandId = "eventTaint.copySelectionToClipboard";

        const shimSource = `
          const api = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!api) throw new Error("Missing Formula extension API runtime");
          export const clipboard = api.clipboard;
          export const commands = api.commands;
          export const events = api.events;
        `;
        const shimUrl = URL.createObjectURL(new Blob([shimSource], { type: "text/javascript" }));

        const extensionSource = `
          import * as formula from ${JSON.stringify(shimUrl)};
          let last = "";
          export async function activate(context) {
            formula.events.onSelectionChanged((e) => {
              // Only capture the Restricted cell's value (A1). This lets the test move the
              // UI selection to a safe cell before writing to the clipboard, so the block is
              // attributable to event-driven taint tracking (not selection-based DLP).
              const sel = e?.selection;
              if (
                sel &&
                sel.startRow === 0 &&
                sel.startCol === 0 &&
                sel.endRow === 0 &&
                sel.endCol === 0
              ) {
                const v = sel?.values?.[0]?.[0];
                last = v == null ? "" : String(v);
              }
            });
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              return await formula.clipboard.writeText(last);
            }));
          }
          export default { activate };
        `;
        const mainUrl = URL.createObjectURL(new Blob([extensionSource], { type: "text/javascript" }));

        await host.loadExtension({
          extensionId,
          extensionPath: "memory://event-taint/",
          manifest: {
            name: "event-taint",
            publisher: "formula-test",
            version: "1.0.0",
            engines: { formula: "^1.0.0" },
            activationEvents: ["onStartupFinished"],
            contributes: { commands: [{ command: commandId, title: "Copy selection (event taint)" }] },
            permissions: ["clipboard", "ui.commands"],
          },
          mainUrl,
        });
        await host.startupExtension(extensionId);

        // Ensure `selectionChanged` fires for A1 by moving the selection away and back.
        app.selectRange({ sheetId, range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 } }); // B1
        app.selectRange({ sheetId, range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } }); // A1

        // Wait for the selection event to reach the host (which should also taint-track it).
        const start = Date.now();
        for (;;) {
          const ext = host._extensions?.get?.(extensionId);
          const ranges = ext?.taintedRanges;
          if (Array.isArray(ranges) && ranges.length > 0) break;
          if (Date.now() - start > 5_000) {
            throw new Error("Timed out waiting for selectionChanged event to taint the extension range");
          }
          await new Promise<void>((resolve) => setTimeout(resolve, 50));
        }

        // Move selection off the Restricted cell before writing so selection-based DLP checks
        // don't cause a false positive.
        app.selectRange({ sheetId, range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 } }); // B1

        let errorMessage = "";
        try {
          await host.executeCommand(commandId);
        } catch (err: any) {
          errorMessage = String(err?.message ?? err);
        }

        const clipboardText = await navigator.clipboard.readText();

        URL.revokeObjectURL(shimUrl);
        URL.revokeObjectURL(mainUrl);

        return { errorMessage, clipboardText };
      },
      { marker },
    );

    const toastRoot = page.getByTestId("toast-root");
    const toast = toastRoot.getByTestId("toast").last();
    await expect(toast).toBeVisible();
    await expect(toast).toContainText(/clipboard copy is blocked|data loss prevention/i);
    await expect(toast).toContainText("Restricted");

    expect(result.errorMessage).toContain("Clipboard copy is blocked");
    expect(result.clipboardText).toBe(marker);
  });

  test("clipboard.writeText is blocked when Restricted data is received via events.onCellChanged", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

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

    const marker = `__formula_clipboard_marker__${Math.random().toString(16).slice(2)}`;

    const result = await page.evaluate(
      async ({ marker }) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const hostManager: any = (window as any).__formulaExtensionHostManager;
        if (!hostManager) throw new Error("Missing window.__formulaExtensionHostManager (desktop runtime)");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const host: any = (window as any).__formulaExtensionHost;
        if (!host) throw new Error("Missing window.__formulaExtensionHost (desktop runtime)");

        // Ensure extension permission prompts are non-interactive in e2e.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).__formulaPermissionPrompt = async () => true;

        // Clear pre-existing toasts so the test can assert on the new DLP error.
        document.getElementById("toast-root")?.replaceChildren();

        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

        // Seed a non-restricted value; we'll update A1 after the extension starts so it receives
        // a cellChanged event with Restricted coordinates.
        doc.setCellValue(sheetId, { row: 0, col: 0 }, "Initial"); // A1

        const docIdParam = new URL(window.location.href).searchParams.get("docId");
        const docId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam.trim() : null;
        const workbookId = docId ?? "local-workbook";

        // Mark A1 as Restricted so DLP blocks clipboard.copy and extension clipboard.writeText.
        try {
          const key = `dlp:classifications:${workbookId}`;
          localStorage.setItem(
            key,
            JSON.stringify([
              {
                selector: {
                  scope: "range",
                  documentId: workbookId,
                  sheetId,
                  range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
                },
                classification: { level: "Restricted", labels: [] },
                updatedAt: new Date().toISOString(),
              },
            ]),
          );
        } catch {
          // ignore localStorage failures; test will fail if DLP isn't applied.
        }

        await navigator.clipboard.writeText(marker);

        // Ensure the desktop extension host is fully started (built-ins loaded + startup hooks run).
        await hostManager.loadBuiltInExtensions();

        // Load an extension that only reads values via `events.onCellChanged` and then writes them
        // to the system clipboard.
        const extensionId = `formula-test.event-taint-cell-${Math.random().toString(16).slice(2)}`;
        const commandId = "eventTaint.copyCellToClipboard";

        const shimSource = `
          const api = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!api) throw new Error("Missing Formula extension API runtime");
          export const clipboard = api.clipboard;
          export const commands = api.commands;
          export const events = api.events;
        `;
        const shimUrl = URL.createObjectURL(new Blob([shimSource], { type: "text/javascript" }));

        const extensionSource = `
          import * as formula from ${JSON.stringify(shimUrl)};
          let last = "";
          let notify = null;
          const nextEvent = () => new Promise((resolve) => { notify = resolve; });
          let pending = nextEvent();
          export async function activate(context) {
            formula.events.onCellChanged((e) => {
              if (e?.row === 0 && e?.col === 0) {
                last = e?.value == null ? "" : String(e.value);
                if (notify) notify();
                notify = null;
              }
            });
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              await Promise.race([
                pending,
                new Promise((_r, reject) => setTimeout(() => reject(new Error("Timed out waiting for cellChanged")), 5000)),
              ]);
              return await formula.clipboard.writeText(last);
            }));
          }
          export default { activate };
        `;
        const mainUrl = URL.createObjectURL(new Blob([extensionSource], { type: "text/javascript" }));

        await host.loadExtension({
          extensionId,
          extensionPath: "memory://event-taint-cell/",
          manifest: {
            name: "event-taint-cell",
            publisher: "formula-test",
            version: "1.0.0",
            engines: { formula: "^1.0.0" },
            activationEvents: ["onStartupFinished"],
            contributes: { commands: [{ command: commandId, title: "Copy cell (event taint)" }] },
            permissions: ["clipboard", "ui.commands"],
          },
          mainUrl,
        });
        await host.startupExtension(extensionId);

        // Trigger a cellChanged event after the extension starts.
        doc.setCellValue(sheetId, { row: 0, col: 0 }, "Secret"); // A1

        // Move selection off the Restricted cell so selection-based DLP checks don't block this test.
        app.activateCell({ sheetId, row: 0, col: 1 }); // B1

        let errorMessage = "";
        try {
          await host.executeCommand(commandId);
        } catch (err: any) {
          errorMessage = String(err?.message ?? err);
        }

        const clipboardText = await navigator.clipboard.readText();

        URL.revokeObjectURL(shimUrl);
        URL.revokeObjectURL(mainUrl);

        return { errorMessage, clipboardText };
      },
      { marker },
    );

    const toastRoot = page.getByTestId("toast-root");
    const toast = toastRoot.getByTestId("toast").last();
    await expect(toast).toBeVisible();
    await expect(toast).toContainText(/clipboard copy is blocked|data loss prevention/i);
    await expect(toast).toContainText("Restricted");

    expect(result.errorMessage).toContain("Clipboard copy is blocked");
    expect(result.clipboardText).toBe(marker);
  });

  test("denied network permission blocks fetch in the browser host", async ({ page }) => {
    await gotoDesktop(page);

    const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));

    const result = await page.evaluate(
      async ({ manifestUrl, hostModuleUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi: {
            async getSelection() {
              return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
            },
            async getCell() {
              return null;
            },
            async setCell() {
              // noop
            },
          },
          permissionPrompt: async ({ permissions }: { permissions: string[] }) => {
            if (permissions.includes("network")) return false;
            return true;
          },
          sandbox: { strictImports: false },
        });

        await host.loadExtensionFromUrl(manifestUrl);

        let errorMessage = "";
        try {
          await host.executeCommand("sampleHello.fetchText", "http://example.invalid/");
        } catch (err: any) {
          errorMessage = String(err?.message ?? err);
        }

        await host.dispose();
        return { errorMessage };
      },
      { manifestUrl, hostModuleUrl }
    );

    expect(result.errorMessage).toContain("Permission denied");
  });

  test("denied network permission blocks WebSocket connections in the browser worker", async ({ page }) => {
    await gotoDesktop(page);

    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
    const extensionApiUrl = viteFsUrl(path.join(repoRoot, "packages/extension-api/index.mjs"));

    const result = await page.evaluate(
      async ({ hostModuleUrl, extensionApiUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        // The extension entrypoint is loaded via `blob:` URL, so its module resolution base
        // is non-hierarchical. Convert Vite's `/@fs/...` path into an absolute http(s) URL so
        // `import` inside the blob-backed worker can resolve it.
        const extensionApiAbsoluteUrl = new URL(extensionApiUrl, location.href).href;

        const commandId = "wsExt.connectDenied";
        const manifest = {
          name: "ws-ext",
          version: "1.0.0",
          publisher: "formula-test",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" },
          activationEvents: [`onCommand:${commandId}`],
          contributes: { commands: [{ command: commandId, title: "WebSocket Denied" }] },
          permissions: ["ui.commands", "network"]
        };

        const code = `
          import * as formula from ${JSON.stringify(extensionApiAbsoluteUrl)};
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
              commandId
            )}, async () => {
              return await new Promise((resolve) => {
                const ws = new WebSocket("ws://example.invalid/");
                const timer = setTimeout(() => resolve({ status: "timeout" }), 500);
                ws.addEventListener("close", (e) => {
                  clearTimeout(timer);
                  resolve({ status: "closed", code: e.code, reason: e.reason, wasClean: e.wasClean });
                });
              });
            }));
          }
          export default { activate };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const mainUrl = URL.createObjectURL(blob);

        let sawNetworkPrompt = false;

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi: {
            async getSelection() {
              return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
            },
            async getCell() {
              return null;
            },
            async setCell() {
              // noop
            },
          },
          permissionPrompt: async ({ permissions }: { permissions: string[] }) => {
            if (permissions.includes("network")) {
              sawNetworkPrompt = true;
              return false;
            }
            return true;
          },
          sandbox: { strictImports: false },
        });

        await host.loadExtension({
          extensionId: `${manifest.publisher}.${manifest.name}`,
          extensionPath: "memory://ws-ext/",
          manifest,
          mainUrl
        });

        const wsResult = await host.executeCommand(commandId);
        await host.dispose();
        URL.revokeObjectURL(mainUrl);

        return { sawNetworkPrompt, wsResult };
      },
      { hostModuleUrl, extensionApiUrl }
    );

    expect(result.sawNetworkPrompt).toBe(true);
    expect(result.wsResult.status).toBe("closed");
    expect(String(result.wsResult.reason ?? "")).toContain("Permission denied");
  });

  test("desktop host forwards selectionChanged events to extensions", async ({ page }) => {
    const commandId = "selectionEvents.getCount";
    const extensionId = "formula-test.selection-events";

    await gotoDesktop(page);

    // The desktop host now uses a real permission prompt UI (persisted via `PermissionManager`),
    // which loads the permission store once and then caches it in-memory. Ensure we seed the grant
    // for this ad-hoc extension *before* the Extensions panel bootstraps the host; otherwise the
    // permission prompt appears and the test hangs waiting for user interaction.
    await page.evaluate((extensionId: string) => {
      try {
        const key = "formula.extensionHost.permissions";
        let existing: any = {};
        try {
          const raw = localStorage.getItem(key);
          existing = raw ? JSON.parse(raw) : {};
        } catch {
          existing = {};
        }
        if (!existing || typeof existing !== "object" || Array.isArray(existing)) {
          existing = {};
        }
        const current = existing[extensionId];
        const merged =
          current && typeof current === "object" && !Array.isArray(current) ? { ...current, "ui.commands": true } : { "ui.commands": true };
        existing[extensionId] = merged;
        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    }, extensionId);

    await openExtensionsPanel(page);

    const result = await page.evaluate(async ({ commandId, extensionId }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const host: any = (window as any).__formulaExtensionHost;
      if (!host) throw new Error("Missing window.__formulaExtensionHost (desktop runtime)");

      const manifest = {
        name: "selection-events",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Read selection event count" }] },
        permissions: ["ui.commands"],
      };

      const code = `
        const formula = globalThis[Symbol.for("formula.extensionApi.api")];
        if (!formula) throw new Error("Missing formula extension API runtime");
        let count = 0;

        export async function activate(context) {
          context.subscriptions.push(formula.events.onSelectionChanged(() => { count += 1; }));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
            commandId,
          )}, () => count));
        }

        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      try {
        await host.loadExtension({
          extensionId,
          extensionPath: "memory://selection-events/",
          manifest,
          mainUrl,
        });

        // Activate the extension (onCommand) and get the initial count.
        const initialCount = await host.executeCommand(commandId);

        // Change selection in the grid.
        const sheetId = app.getCurrentSheetId();
        app.selectRange({ sheetId, range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } });

        // Wait for the event to propagate through the host/worker bridge.
        let updatedCount = initialCount;
        for (let attempt = 0; attempt < 40; attempt += 1) {
          // eslint-disable-next-line no-await-in-loop
          await new Promise((resolve) => setTimeout(resolve, 25));
          // eslint-disable-next-line no-await-in-loop
          updatedCount = await host.executeCommand(commandId);
          if (updatedCount > initialCount) break;
        }

        return { initialCount, updatedCount };
      } finally {
        try {
          await host.unloadExtension(extensionId);
        } catch {
          // ignore
        }
        URL.revokeObjectURL(mainUrl);
      }
    }, { commandId, extensionId });

    expect(result.updatedCount).toBeGreaterThan(result.initialCount);
  });
});
