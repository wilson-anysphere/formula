import { expect, test } from "@playwright/test";

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
});

