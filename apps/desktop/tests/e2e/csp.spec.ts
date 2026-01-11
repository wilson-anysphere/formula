import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

test.describe("Content Security Policy (Tauri parity)", () => {
  test("startup has no CSP violations and allows WASM-in-worker execution", async ({ page }) => {
    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toBeVisible();

    const { mainThreadAnswer, workerAnswer } = await page.evaluate(async () => {
      const wasmBytes = new Uint8Array([
        0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00, 0x01, 0x05, 0x01, 0x60, 0x00, 0x01, 0x7f,
        0x03, 0x02, 0x01, 0x00, 0x07, 0x0a, 0x01, 0x06, 0x61, 0x6e, 0x73, 0x77, 0x65, 0x72, 0x00,
        0x00, 0x0a, 0x06, 0x01, 0x04, 0x00, 0x41, 0x2a, 0x0b
      ]);

      const { instance } = await WebAssembly.instantiate(wasmBytes);
      const mainThreadAnswer = (instance.exports as any).answer() as number;

      const workerAnswer = await new Promise<number>((resolve, reject) => {
        const code = `
          const wasmBytes = new Uint8Array(${JSON.stringify(Array.from(wasmBytes))});
          self.onmessage = async () => {
            try {
              const { instance } = await WebAssembly.instantiate(wasmBytes);
              self.postMessage(instance.exports.answer());
            } catch (err) {
              self.postMessage({ error: String(err) });
            }
          };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const url = URL.createObjectURL(blob);
        const worker = new Worker(url, { type: "module" });

        const cleanup = () => {
          worker.terminate();
          URL.revokeObjectURL(url);
        };

        worker.onmessage = (event) => {
          const payload = event.data as any;
          if (payload && typeof payload === "object" && "error" in payload) {
            cleanup();
            reject(new Error(String(payload.error)));
            return;
          }
          cleanup();
          resolve(payload as number);
        };

        worker.onerror = (event) => {
          cleanup();
          reject(new Error((event as ErrorEvent).message || "worker error"));
        };

        worker.postMessage(null);
      });

      return { mainThreadAnswer, workerAnswer };
    });

    expect(mainThreadAnswer).toBe(42);
    expect(workerAnswer).toBe(42);

    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });

  test("BrowserExtensionHost can run an extension in a Worker without CSP violations", async ({ page }) => {
    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toBeVisible();

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
          permissionPrompt: async () => true
        });

        await host.loadExtensionFromUrl(manifestUrl);
        const sum = await host.executeCommand("sampleHello.sumSelection");
        const a3 = app.getCellValueA1("A3");
        await host.dispose();

        return { sum, a3 };
      },
      { manifestUrl, hostModuleUrl }
    );

    expect(result.sum).toBe(10);
    expect(result.a3).toBe("10");

    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });
});
