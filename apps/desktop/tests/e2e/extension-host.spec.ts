import { expect, test } from "@playwright/test";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop } from "./helpers";

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
});
