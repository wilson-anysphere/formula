import test from "node:test";
import assert from "node:assert/strict";

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

test("clipboard provider", async (t) => {
  await t.test("web: read returns html/text when available", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain", "text/html"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return new Blob(["hello"], { type });
                      case "text/html":
                        return new Blob(["<b>hi</b>"], { type });
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
        assert.deepEqual(content, { text: "hello", html: "<b>hi</b>" });
      }
    );
  });

  await t.test("web: read falls back to readText when rich read throws", async () => {
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

  await t.test("web: read ignores oversized image/png blobs", async () => {
    const large = new Blob([new ArrayBuffer(11 * 1024 * 1024)], { type: "image/png" });

    // Ensure we don't accidentally allocate another large ArrayBuffer by calling
    // `Blob#arrayBuffer()`.
    // @ts-ignore
    large.arrayBuffer = async () => {
      throw new Error("should not call arrayBuffer() for oversized images");
    };

    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["image/png", "text/plain"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "image/png":
                        return large;
                      case "text/plain":
                        return new Blob(["hello"], { type });
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
        assert.deepEqual(content, { text: "hello" });
      }
    );
  });

  await t.test("web: read returns small image/png blobs", async () => {
    const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["image/png"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "image/png":
                        return new Blob([bytes], { type });
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

        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { imagePng: bytes });
      }
    );
  });

  await t.test("web: write uses navigator.clipboard.write with text/plain + text/html", async () => {
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
        await provider.write({ text: "plain", html: "<p>hello</p>" });

        assert.equal(writes.length, 1);
        assert.equal(writeTextCalls.length, 0);

        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["text/html", "text/plain"].sort());

        assert.equal(item.data["text/plain"].type, "text/plain");
        assert.equal(await item.data["text/plain"].text(), "plain");

        assert.equal(item.data["text/html"].type, "text/html");
        assert.equal(await item.data["text/html"].text(), "<p>hello</p>");
      }
    );
  });

  await t.test("web: write falls back to writeText when rich write throws", async () => {
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
        await provider.write({ text: "plain", html: "<p>x</p>" });
        assert.deepEqual(writeTextCalls, ["plain"]);
      }
    );
  });

  await t.test("web: provider tolerates missing Clipboard APIs", async () => {
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

  await t.test("tauri: read prefers core.invoke('clipboard_read')", async () => {
    /** @type {any[]} */
    const invokeCalls = [];
    /** @type {number} */
    let webReadCalls = 0;

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              invokeCalls.push(cmd);
              assert.equal(cmd, "clipboard_read");
              return { text: "tauri text", html: "tauri html", rtf: "tauri rtf", pngBase64: "CQgH" };
            },
          },
          clipboard: {
            async readText() {
              throw new Error("should not read text when clipboard_read succeeds");
            },
          },
        },
        navigator: {
          clipboard: {
            async read() {
              webReadCalls += 1;
              throw new Error("should not call web clipboard when clipboard_read succeeds");
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, {
          text: "tauri text",
          html: "tauri html",
          rtf: "tauri rtf",
          pngBase64: "CQgH",
        });
        assert.equal(webReadCalls, 0);
        assert.deepEqual(invokeCalls, ["clipboard_read"]);
      }
    );
  });

  await t.test("tauri: read preserves pngBase64 when native result only includes image_png_base64", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { image_png_base64: "CQgH" };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: undefined, pngBase64: "CQgH" });
      }
    );
  });

  await t.test("tauri: read falls back to web clipboard before tauri clipboard.readText", async () => {
    /** @type {number} */
    let tauriReadTextCalls = 0;

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return {};
            },
          },
          clipboard: {
            async readText() {
              tauriReadTextCalls += 1;
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
              return "web text";
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: "web text" });
        assert.equal(tauriReadTextCalls, 0);
      }
    );
  });

  await t.test("tauri: read falls back to tauri clipboard.readText when web clipboard yields no content", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              if (cmd === "clipboard_read" || cmd === "read_clipboard") {
                throw new Error("command not found");
              }
              throw new Error(`Unexpected command: ${cmd}`);
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

  await t.test("tauri: read falls back to core.invoke('read_clipboard') when clipboard_read is unavailable", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              if (cmd === "clipboard_read") {
                throw new Error("command not found");
              }
              if (cmd === "read_clipboard") {
                return { text: "legacy text", html: "legacy html", rtf: "legacy rtf", image_png_base64: "CQgH" };
              }
              throw new Error(`Unexpected command: ${cmd}`);
            },
          },
          clipboard: {
            async readText() {
              throw new Error("should not call readText when read_clipboard succeeds");
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
        assert.deepEqual(content, {
          text: "legacy text",
          html: "legacy html",
          rtf: "legacy rtf",
          pngBase64: "CQgH",
        });
      }
    );
  });

  await t.test("tauri: provider tolerates missing core.invoke and clipboard APIs", async () => {
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

  await t.test("tauri: write invokes core.invoke('clipboard_write') and then attempts ClipboardItem write", async () => {
    /** @type {any[]} */
    const invokeCalls = [];
    /** @type {any[]} */
    const writes = [];

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
        __TAURI__: {
          core: {
            async invoke(cmd, args) {
              invokeCalls.push([cmd, args]);
            },
          },
          clipboard: {
            async writeText() {
              throw new Error("should not call legacy writeText when clipboard_write succeeds");
            },
          },
        },
        ClipboardItem: MockClipboardItem,
        navigator: {
          clipboard: {
            async write(items) {
              writes.push(items);
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain", html: "<p>hello</p>", rtf: "{\\\\rtf1 hello}", pngBase64: "CQgH" });

        assert.equal(invokeCalls.length, 1);
        assert.equal(invokeCalls[0][0], "clipboard_write");
        assert.deepEqual(invokeCalls[0][1], {
          payload: {
            text: "plain",
            html: "<p>hello</p>",
            rtf: "{\\\\rtf1 hello}",
            pngBase64: "CQgH",
          },
        });

        assert.equal(writes.length, 1);
        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        // The WebView ClipboardItem write is best-effort and currently only includes
        // HTML + plain text. Rich formats like RTF are written via the native
        // `clipboard_write` command above.
        assert.deepEqual(keys, ["text/html", "text/plain"].sort());
      }
    );
  });

  await t.test("tauri: write falls back to legacy __TAURI__.clipboard.writeText when clipboard_write throws", async () => {
    /** @type {string[]} */
    const writeTextCalls = [];

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              if (cmd === "clipboard_write" || cmd === "write_clipboard") {
                throw new Error("command not found");
              }
              throw new Error(`Unexpected command: ${cmd}`);
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

  await t.test("tauri: write falls back to core.invoke('write_clipboard') when clipboard_write is unavailable", async () => {
    /** @type {any[]} */
    const invokeCalls = [];

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd, args) {
              invokeCalls.push([cmd, args]);
              if (cmd === "clipboard_write") {
                throw new Error("command not found");
              }
              if (cmd === "write_clipboard") {
                return;
              }
              throw new Error(`Unexpected command: ${cmd}`);
            },
          },
          clipboard: {
            async writeText() {
              throw new Error("should not call legacy writeText when write_clipboard succeeds");
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "hello", html: "<p>x</p>", pngBase64: "CQgH" });

        assert.equal(invokeCalls.length, 2);
        assert.equal(invokeCalls[0][0], "clipboard_write");
        assert.equal(invokeCalls[1][0], "write_clipboard");
        assert.deepEqual(invokeCalls[1][1], {
          text: "hello",
          html: "<p>x</p>",
          rtf: undefined,
          image_png_base64: "CQgH",
        });
      }
    );
  });
});
