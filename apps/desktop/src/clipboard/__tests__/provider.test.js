import test from "node:test";
import assert from "node:assert/strict";
import { Buffer as NodeBuffer } from "node:buffer";

import { createClipboardProvider } from "../platform/provider.js";

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

test("clipboard provider: rich MIME handling", async (t) => {
  await t.test("web read returns html/text/rtf/image when available", async () => {
    const imageBytes = Uint8Array.from([1, 2, 3, 4]);

    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain", "text/html", "text/rtf", "image/png"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return new Blob(["hello"], { type });
                      case "text/html":
                        return new Blob(["<b>hi</b>"], { type });
                      case "text/rtf":
                        return new Blob(["{\\\\rtf1 hi}"], { type });
                      case "image/png":
                        return new Blob([imageBytes], { type });
                      default:
                        throw new Error(`Unexpected clipboard type: ${type}`);
                    }
                  },
                },
              ];
            },
            async readText() {
              return "fallback";
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();

        assert.equal(content.text, "hello");
        assert.equal(content.html, "<b>hi</b>");
        assert.equal(content.rtf, "{\\\\rtf1 hi}");
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(Array.from(content.imagePng), Array.from(imageBytes));
      }
    );
  });

  await t.test("web write uses navigator.clipboard.write with all provided MIME keys", async () => {
    const imageBytes = Uint8Array.from([5, 6, 7]);

    /** @type {any[]} */
    const writes = [];
    /** @type {string[]} */
    const writeTextCalls = [];

    class MockClipboardItem {
      /**
       * @param {Record<string, Blob>} data
       */
      constructor(data) {
        this.data = data;
      }
    }

    await withGlobals(
      {
        __TAURI__: undefined,
        ClipboardItem: MockClipboardItem,
        navigator: {
          clipboard: {
            async write(items) {
              writes.push(items);
            },
            async writeText(text) {
              writeTextCalls.push(text);
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({
          text: "plain",
          html: "<p>hello</p>",
          rtf: "{\\\\rtf1 hello}",
          imagePng: imageBytes,
        });

        assert.equal(writes.length, 1);
        assert.equal(writeTextCalls.length, 0);

        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["image/png", "text/html", "text/plain", "text/rtf"].sort());

        assert.equal(item.data["text/plain"].type, "text/plain");
        assert.equal(await item.data["text/plain"].text(), "plain");

        assert.equal(item.data["text/html"].type, "text/html");
        assert.equal(await item.data["text/html"].text(), "<p>hello</p>");

        assert.equal(item.data["text/rtf"].type, "text/rtf");
        assert.equal(await item.data["text/rtf"].text(), "{\\\\rtf1 hello}");

        assert.equal(item.data["image/png"].type, "image/png");
        const ab = await item.data["image/png"].arrayBuffer();
        assert.deepEqual(Array.from(new Uint8Array(ab)), Array.from(imageBytes));
      }
    );
  });

  await t.test("web read falls back to readText when rich read throws", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              throw new Error("permission denied");
            },
            async readText() {
              return "fallback text";
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: "fallback text" });
      }
    );
  });

  await t.test("web write falls back to writeText when rich write throws", async () => {
    /** @type {string[]} */
    const writeTextCalls = [];

    class MockClipboardItem {
      constructor() {}
    }

    await withGlobals(
      {
        __TAURI__: undefined,
        ClipboardItem: MockClipboardItem,
        navigator: {
          clipboard: {
            async write() {
              throw new Error("write not allowed");
            },
            async writeText(text) {
              writeTextCalls.push(text);
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain" });
        assert.deepEqual(writeTextCalls, ["plain"]);
      }
    );
  });

  await t.test("web provider tolerates missing Clipboard APIs", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {},
        ClipboardItem: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: undefined });

        await provider.write({ text: "plain" });
      }
    );
  });

  await t.test("tauri provider merges invoke('read_clipboard') results without clobbering web fields", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            /**
             * @param {string} cmd
             */
            async invoke(cmd) {
              assert.equal(cmd, "read_clipboard");
              return {
                text: "tauri text",
                html: "tauri html",
                rtf: "tauri rtf",
                image_png_base64: "CQgH", // [9, 8, 7]
              };
            },
          },
          clipboard: {},
        },
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain", "text/html"],
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return new Blob(["web text"], { type });
                      case "text/html":
                        return new Blob(["<div>web</div>"], { type });
                      default:
                        throw new Error(`Unexpected clipboard type: ${type}`);
                    }
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

        // Web clipboard values win.
        assert.equal(content.text, "web text");
        assert.equal(content.html, "<div>web</div>");

        // Tauri invoke fills in missing rich formats.
        assert.equal(content.rtf, "tauri rtf");
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(Array.from(content.imagePng), [9, 8, 7]);
      }
    );
  });

  await t.test("tauri provider decodes image base64 without Buffer (atob fallback)", async () => {
    await withGlobals(
      {
        Buffer: undefined,
        atob: (base64) => NodeBuffer.from(base64, "base64").toString("binary"),
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "read_clipboard");
              return { image_png_base64: "CQgH" }; // [9, 8, 7]
            },
          },
          clipboard: {},
        },
        navigator: { clipboard: {} },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(Array.from(content.imagePng), [9, 8, 7]);
      }
    );
  });

  await t.test("tauri provider falls back gracefully when invoke throws", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke() {
              throw new Error("command not found");
            },
          },
          clipboard: {
            async readText() {
              return "tauri text";
            },
          },
        },
        navigator: {
          clipboard: {
            async read() {
              throw new Error("permission denied");
            },
            async readText() {
              throw new Error("permission denied");
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: "tauri text" });
      }
    );
  });

  await t.test("tauri provider prefers tauri readText over web readText fallback", async () => {
    /** @type {number} */
    let webReadTextCalls = 0;

    await withGlobals(
      {
        __TAURI__: {
          clipboard: {
            async readText() {
              return "tauri text";
            },
          },
        },
        navigator: {
          clipboard: {
            async read() {
              throw new Error("permission denied");
            },
            async readText() {
              webReadTextCalls += 1;
              return "web text";
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: "tauri text" });
        assert.equal(webReadTextCalls, 0);
      }
    );
  });

  await t.test("tauri provider tolerates missing core.invoke and clipboard APIs", async () => {
    await withGlobals(
      {
        __TAURI__: {},
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: undefined });

        await provider.write({ text: "plain" });
      }
    );
  });

  await t.test("tauri provider write calls invoke('write_clipboard') for rich payloads", async () => {
    const imageBytes = Uint8Array.from([9, 8, 7]);

    /** @type {any[]} */
    const invokeCalls = [];
    /** @type {string[]} */
    const writeTextCalls = [];

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd, args) {
              invokeCalls.push([cmd, args]);
            },
          },
          clipboard: {
            async writeText(text) {
              writeTextCalls.push(text);
            },
          },
        },
        navigator: {
          clipboard: {
            async write() {
              throw new Error("web clipboard should not be used when invoke succeeds");
            },
            async writeText() {
              throw new Error("web clipboard should not be used when invoke succeeds");
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({
          text: "plain",
          html: "<p>hello</p>",
          imagePng: imageBytes,
        });

        assert.deepEqual(writeTextCalls, ["plain"]);
        assert.equal(invokeCalls.length, 1);

        const [cmd, args] = invokeCalls[0];
        assert.equal(cmd, "write_clipboard");
        assert.equal(args.text, "plain");
        assert.equal(args.html, "<p>hello</p>");
        assert.equal(args.rtf, undefined);
        assert.equal(args.image_png_base64, "CQgH");
      }
    );
  });

  await t.test("tauri provider encodes image base64 without Buffer (btoa fallback)", async () => {
    const imageBytes = Uint8Array.from([9, 8, 7]);

    /** @type {any[]} */
    const invokeCalls = [];

    await withGlobals(
      {
        Buffer: undefined,
        btoa: (binary) => NodeBuffer.from(binary, "binary").toString("base64"),
        __TAURI__: {
          core: {
            async invoke(cmd, args) {
              invokeCalls.push([cmd, args]);
            },
          },
          clipboard: {},
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain", imagePng: imageBytes });

        assert.equal(invokeCalls.length, 1);
        const [cmd, args] = invokeCalls[0];
        assert.equal(cmd, "write_clipboard");
        assert.equal(args.image_png_base64, "CQgH");
      }
    );
  });

  await t.test("tauri provider write falls back when invoke throws", async () => {
    /** @type {string[]} */
    const writeTextCalls = [];

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke() {
              throw new Error("command not found");
            },
          },
          clipboard: {
            async writeText(text) {
              writeTextCalls.push(text);
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "hello", html: "<p>x</p>" });
        assert.deepEqual(writeTextCalls, ["hello"]);
      }
    );
  });
});
