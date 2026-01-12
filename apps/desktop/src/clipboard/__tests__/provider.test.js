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

  await t.test("web: read recognizes rich mime types with charset parameters", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain;charset=utf-8", "text/html;charset=utf-8", "text/rtf;charset=utf-8"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain;charset=utf-8":
                        return new Blob(["hello"], { type });
                      case "text/html;charset=utf-8":
                        return new Blob(["<b>hi</b>"], { type });
                      case "text/rtf;charset=utf-8":
                        return new Blob(["{\\\\rtf1 hello}"], { type });
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
        assert.deepEqual(content, { text: "hello", html: "<b>hi</b>", rtf: "{\\\\rtf1 hello}" });
      }
    );
  });

  await t.test("web: read recognizes rich mime types case-insensitively", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["TEXT/PLAIN", "TEXT/HTML", "TEXT/RTF"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "TEXT/PLAIN":
                        return new Blob(["hello"], { type });
                      case "TEXT/HTML":
                        return new Blob(["<b>hi</b>"], { type });
                      case "TEXT/RTF":
                        return new Blob(["{\\\\rtf1 hello}"], { type });
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
        assert.deepEqual(content, { text: "hello", html: "<b>hi</b>", rtf: "{\\\\rtf1 hello}" });
      }
    );
  });

  await t.test("web: read recognizes application/x-rtf", async () => {
    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain", "application/x-rtf"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return new Blob(["hello"], { type });
                      case "application/x-rtf":
                        return new Blob(["{\\\\rtf1 hello}"], { type });
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
        assert.deepEqual(content, { text: "hello", rtf: "{\\\\rtf1 hello}" });
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

  await t.test("web: read ignores oversized text/html blobs (keeps plain text)", async () => {
    const largeHtml = new Blob([new ArrayBuffer(3 * 1024 * 1024)], { type: "text/html" });

    // Ensure we don't attempt to materialize a huge string via `Blob#text()`.
    // @ts-ignore
    largeHtml.text = async () => {
      throw new Error("should not call text() for oversized html blobs");
    };

    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/html", "text/plain"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/html":
                        return largeHtml;
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

  await t.test("web: read ignores oversized text/plain blobs (no readText fallback)", async () => {
    const largeText = new Blob([new ArrayBuffer(3 * 1024 * 1024)], { type: "text/plain" });
    // Ensure we don't attempt to materialize a huge string via `Blob#text()`.
    // @ts-ignore
    largeText.text = async () => {
      throw new Error("should not call text() for oversized plain text blobs");
    };

    await withGlobals(
      {
        __TAURI__: undefined,
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return largeText;
                      default:
                        throw new Error(`Unexpected clipboard type: ${type}`);
                    }
                  },
                },
              ];
            },
            async readText() {
              throw new Error("should not call readText() when oversized plain text is detected");
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, {});
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

  await t.test("web: write includes text/rtf when provided", async () => {
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
        __TAURI__: undefined,
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
        await provider.write({ text: "plain", html: "<p>hello</p>", rtf: "{\\\\rtf1 hello}" });

        assert.equal(writes.length, 1);
        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["text/html", "text/plain", "text/rtf"].sort());

        assert.equal(item.data["text/rtf"].type, "text/rtf");
        assert.equal(await item.data["text/rtf"].text(), "{\\\\rtf1 hello}");
      }
    );
  });

  await t.test("web: write falls back to html/plain when text/rtf ClipboardItem write fails", async () => {
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
              const keys = Object.keys(items?.[0]?.data ?? {});
              if (keys.includes("text/rtf")) {
                throw new Error("unsupported type: text/rtf");
              }
            },
            async writeText(text) {
              writeTextCalls.push(text);
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain", html: "<p>hello</p>", rtf: "{\\\\rtf1 hello}" });

        // First attempt includes RTF, then we retry without it.
        assert.equal(writes.length, 2);
        assert.equal(writeTextCalls.length, 0);

        const firstKeys = Object.keys(writes[0][0].data).sort();
        assert.deepEqual(firstKeys, ["text/html", "text/plain", "text/rtf"].sort());

        const secondKeys = Object.keys(writes[1][0].data).sort();
        assert.deepEqual(secondKeys, ["text/html", "text/plain"].sort());
      }
    );
  });

  await t.test("web: write includes small image/png when provided", async () => {
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

    const imageBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

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
        await provider.write({ text: "plain", html: "<p>hello</p>", imagePng: imageBytes });

        assert.equal(writes.length, 1);
        assert.equal(writeTextCalls.length, 0);

        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["image/png", "text/html", "text/plain"].sort());

        assert.equal(item.data["image/png"].type, "image/png");
        assert.equal(item.data["image/png"].size, imageBytes.byteLength);
      }
    );
  });

  await t.test("web: write includes small image/png when provided via legacy pngBase64", async () => {
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

    // Minimal PNG signature bytes.
    const pngBase64 = "iVBORw==";

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
        await provider.write({ text: "plain", html: "<p>hello</p>", pngBase64 });

        assert.equal(writes.length, 1);
        assert.equal(writeTextCalls.length, 0);

        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["image/png", "text/html", "text/plain"].sort());

        assert.equal(item.data["image/png"].type, "image/png");
        assert.equal(item.data["image/png"].size, 4);
      }
    );
  });

  await t.test("web: write omits oversized image/png blobs but still writes html/text", async () => {
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

    const largeImage = new Blob([new ArrayBuffer(11 * 1024 * 1024)], { type: "image/png" });

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
        await provider.write({ text: "plain", html: "<p>hello</p>", imagePng: largeImage });

        assert.equal(writes.length, 1);
        assert.equal(writeTextCalls.length, 0);

        assert.equal(writes[0].length, 1);
        const item = writes[0][0];
        assert.ok(item instanceof MockClipboardItem);

        const keys = Object.keys(item.data).sort();
        assert.deepEqual(keys, ["text/html", "text/plain"].sort());
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
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, {
          text: "tauri text",
          html: "tauri html",
          rtf: "tauri rtf",
          imagePng: new Uint8Array([0x09, 0x08, 0x07]),
        });
        assert.equal(webReadCalls, 0);
        assert.deepEqual(invokeCalls, ["clipboard_read"]);
      }
    );
  });

  await t.test("tauri: read decodes image_png_base64 into imagePng bytes", async () => {
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
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { text: undefined, imagePng: new Uint8Array([0x09, 0x08, 0x07]) });
      }
    );
  });

  await t.test("tauri: read decodes png_base64 into imagePng bytes", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { png_base64: "CQgH" };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { text: undefined, imagePng: new Uint8Array([0x09, 0x08, 0x07]) });
      }
    );
  });

  await t.test("tauri: read decodes data URL pngBase64 into imagePng bytes", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { pngBase64: "data:image/png;base64,CQgH" };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { text: undefined, imagePng: new Uint8Array([0x09, 0x08, 0x07]) });
      }
    );
  });

  await t.test("tauri: read decodes data URL pngBase64 case-insensitively (with whitespace)", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { pngBase64: " \n DATA:image/png;base64,CQgH \n" };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { text: undefined, imagePng: new Uint8Array([0x09, 0x08, 0x07]) });
      }
    );
  });

  await t.test("tauri: read decodes data URL pngBase64 when atob is unavailable (Buffer fallback)", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { pngBase64: "DATA:image/png;base64,CQgH" };
            },
          },
        },
        navigator: undefined,
        atob: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, { text: undefined, imagePng: new Uint8Array([0x09, 0x08, 0x07]) });
      }
    );
  });

  await t.test("tauri: read drops oversized pngBase64 payloads", async () => {
    const largeBase64 = "A".repeat(14 * 1024 * 1024);

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { text: "tauri text", pngBase64: largeBase64 };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: "tauri text" });
      }
    );
  });

  await t.test("tauri: read ignores empty data URL pngBase64 payloads", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { pngBase64: "data:image/png;base64," };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: undefined });
      }
    );
  });

  await t.test("tauri: read ignores malformed data URL pngBase64 payloads without a comma", async () => {
    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd) {
              assert.equal(cmd, "clipboard_read");
              return { pngBase64: "data:image/png;base64" };
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, { text: undefined });
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

  await t.test("tauri: read avoids tauri clipboard.readText when web clipboard detects oversized text/plain", async () => {
    const largeText = new Blob([new ArrayBuffer(3 * 1024 * 1024)], { type: "text/plain" });
    // Ensure we don't attempt to materialize a huge string via `Blob#text()`.
    // @ts-ignore
    largeText.text = async () => {
      throw new Error("should not call text() for oversized plain text blobs");
    };

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
              throw new Error("should not call tauri clipboard.readText when oversized plain text is detected");
            },
          },
        },
        navigator: {
          clipboard: {
            async read() {
              return [
                {
                  types: ["text/plain"],
                  /**
                   * @param {string} type
                   */
                  async getType(type) {
                    switch (type) {
                      case "text/plain":
                        return largeText;
                      default:
                        throw new Error(`Unexpected clipboard type: ${type}`);
                    }
                  },
                },
              ];
            },
            async readText() {
              throw new Error("should not call readText() when oversized plain text is detected");
            },
          },
        },
      },
      async () => {
        const provider = await createClipboardProvider();
        const content = await provider.read();
        assert.deepEqual(content, {});
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
        assert.ok(content.imagePng instanceof Uint8Array);
        assert.deepEqual(content, {
          text: "legacy text",
          html: "legacy html",
          rtf: "legacy rtf",
          imagePng: new Uint8Array([0x09, 0x08, 0x07]),
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

  await t.test("tauri: write invokes core.invoke('clipboard_write') and does not attempt ClipboardItem write", async () => {
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
        await provider.write({
          text: "plain",
          html: "<p>hello</p>",
          rtf: "{\\\\rtf1 hello}",
          imagePng: new Uint8Array([0x09, 0x08, 0x07]),
        });

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

        // If the native `clipboard_write` command succeeds, avoid the Web Clipboard API
        // write, since it can clobber other rich formats (RTF/image) that were written natively.
        assert.equal(writes.length, 0);
      }
    );
  });

  await t.test("tauri: write accepts legacy pngBase64 when imagePng is omitted", async () => {
    /** @type {any[]} */
    const invokeCalls = [];

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
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain", html: "<p>hello</p>", pngBase64: "data:image/png;base64,CQgH" });

        assert.equal(invokeCalls.length, 1);
        assert.equal(invokeCalls[0][0], "clipboard_write");
        assert.deepEqual(invokeCalls[0][1], {
          payload: {
            text: "plain",
            html: "<p>hello</p>",
            pngBase64: "CQgH",
          },
        });
      }
    );
  });

  await t.test("tauri: write normalizes data URL pngBase64 case-insensitively (with whitespace)", async () => {
    /** @type {any[]} */
    const invokeCalls = [];

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
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();
        await provider.write({ text: "plain", html: "<p>hello</p>", pngBase64: " \n DATA:image/png;base64,CQgH \n" });

        assert.equal(invokeCalls.length, 1);
        assert.equal(invokeCalls[0][0], "clipboard_write");
        assert.deepEqual(invokeCalls[0][1], {
          payload: {
            text: "plain",
            html: "<p>hello</p>",
            pngBase64: "CQgH",
          },
        });
      }
    );
  });

  await t.test("tauri: write attempts ClipboardItem write when native invoke fails and html is present", async () => {
    /** @type {string[]} */
    const writeTextCalls = [];
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
        await provider.write({ text: "plain", html: "<p>hello</p>" });

        // Plain text fallback still happens via the legacy Tauri clipboard API.
        assert.deepEqual(writeTextCalls, ["plain"]);

        // And we still attempt an HTML write via the Web Clipboard API when native rich writes fail.
        assert.equal(writes.length, 1);
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

  await t.test("tauri: write base64-encodes imagePng and falls back to core.invoke('write_clipboard') when clipboard_write is unavailable", async () => {
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
        await provider.write({ text: "hello", html: "<p>x</p>", imagePng: new Uint8Array([0x09, 0x08, 0x07]) });

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

  await t.test("tauri: write omits oversized imagePng payloads (MAX_IMAGE_BYTES guard)", async () => {
    /** @type {any[]} */
    const invokeCalls = [];

    await withGlobals(
      {
        __TAURI__: {
          core: {
            async invoke(cmd, args) {
              invokeCalls.push([cmd, args]);
            },
          },
        },
        navigator: undefined,
      },
      async () => {
        const provider = await createClipboardProvider();

        const oversized = new Uint8Array(11 * 1024 * 1024);
        await provider.write({ text: "plain", imagePng: oversized });

        assert.equal(invokeCalls.length, 1);
        assert.equal(invokeCalls[0][0], "clipboard_write");
        assert.deepEqual(invokeCalls[0][1], { payload: { text: "plain" } });
      }
    );
  });
});
