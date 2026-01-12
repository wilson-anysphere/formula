import test from "node:test";
import assert from "node:assert/strict";

import { createClipboardProvider } from "../src/clipboard/platform/provider.js";

/**
 * @param {Record<string, any>} overrides
 * @param {() => Promise<void> | void} fn
 */
async function withGlobals(overrides, fn) {
  /** @type {Map<string, PropertyDescriptor | undefined>} */
  const originals = new Map();

  for (const [key, value] of Object.entries(overrides)) {
    originals.set(key, Object.getOwnPropertyDescriptor(globalThis, key));
    Object.defineProperty(globalThis, key, {
      value,
      configurable: true,
      writable: true,
      enumerable: true,
    });
  }

  try {
    await fn();
  } finally {
    for (const [key, desc] of originals.entries()) {
      if (desc) {
        Object.defineProperty(globalThis, key, desc);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete globalThis[key];
      }
    }
  }
}

test("tauri: preserves rich WebView formats when filling missing text from tauri clipboard.readText", async () => {
  const rtf = String.raw`{\rtf1\ansi hello}`;
  const imageBytes = new Uint8Array([1, 2, 3, 4, 5]);

  await withGlobals(
    {
      __TAURI__: {
        clipboard: {
          async readText() {
            return "foo";
          },
        },
      },
      navigator: {
        clipboard: {
          async read() {
            return [
              {
                types: ["text/rtf", "image/png"],
                async getType(type) {
                  if (type === "text/rtf") return new Blob([rtf], { type });
                  if (type === "image/png") return new Blob([imageBytes], { type });
                  throw new Error(`unexpected clipboard type: ${type}`);
                },
              },
            ];
          },
        },
      },
    },
    async () => {
      const provider = await createClipboardProvider();
      const content = await provider.read();

      assert.equal(content.text, "foo");
      assert.equal(content.rtf, rtf);
      assert.ok(content.imagePng instanceof Uint8Array);
      assert.deepStrictEqual(content.imagePng, imageBytes);
    },
  );
});

