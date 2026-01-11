import { expect, test } from "@playwright/test";

test("can load wasm-bindgen assets without CSP/wasm errors", async ({ page }) => {
  const wasmErrors: string[] = [];

  page.on("console", (msg) => {
    if (msg.type() !== "error") return;
    const text = msg.text();
    if (/wasm|webassembly|content security policy|csp/i.test(text)) {
      wasmErrors.push(text);
    }
  });

  page.on("pageerror", (err) => {
    const text = err.message ?? String(err);
    if (/wasm|webassembly|content security policy|csp/i.test(text)) {
      wasmErrors.push(text);
    }
  });

  await page.goto("/");

  const result = await page.evaluate(async () => {
    return await new Promise<{ ok: boolean; value?: unknown; error?: string }>((resolve, reject) => {
      const worker = new Worker("/engine/wasm_smoke.worker.js", { type: "module" });
      worker.addEventListener("message", (event) => {
        worker.terminate();
        resolve(event.data);
      });
      worker.addEventListener("error", (event) => {
        worker.terminate();
        reject(new Error(event.message));
      });
      worker.postMessage(null);
    });
  });

  expect(result.ok).toBe(true);
  expect(result.value).toBe(2);
  expect(wasmErrors).toEqual([]);
});

