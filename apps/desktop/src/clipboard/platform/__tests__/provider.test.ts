import { afterEach, describe, expect, it, vi } from "vitest";

import { createClipboardProvider } from "../provider.js";

type ClipboardMocks = {
  readText?: ReturnType<typeof vi.fn>;
  writeText?: ReturnType<typeof vi.fn>;
  read?: ReturnType<typeof vi.fn>;
  write?: ReturnType<typeof vi.fn>;
};

const originalNavigatorDescriptor = Object.getOwnPropertyDescriptor(globalThis, "navigator");
const originalTauri = (globalThis as any).__TAURI__;

function setMockNavigatorClipboard(clipboard: ClipboardMocks) {
  Object.defineProperty(globalThis, "navigator", {
    value: { clipboard },
    configurable: true,
    writable: true,
  });
}

afterEach(() => {
  if (originalNavigatorDescriptor) {
    Object.defineProperty(globalThis, "navigator", originalNavigatorDescriptor);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).navigator;
  }

  if (typeof originalTauri === "undefined") {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).__TAURI__;
  } else {
    (globalThis as any).__TAURI__ = originalTauri;
  }

  vi.restoreAllMocks();
});

describe("clipboard/platform/provider (desktop Tauri multi-format path)", () => {
  it("read() uses __TAURI__.core.invoke('clipboard_read') when available and returns the payload", async () => {
    const invoke = vi.fn().mockResolvedValue({ text: "hello", html: "<p>hello</p>" });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const readText = vi.fn().mockResolvedValue("web-fallback");
    setMockNavigatorClipboard({ readText });

    const provider = await createClipboardProvider();
    const content = await provider.read();

    expect(content).toEqual({ text: "hello", html: "<p>hello</p>" });
    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("clipboard_read");

    // No fallback should be needed.
    expect(readText).not.toHaveBeenCalled();
  });

  it("read() merges a text-only Tauri payload with richer WebView clipboard HTML", async () => {
    const invoke = vi.fn().mockResolvedValue({ text: "native-text" });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const read = vi.fn().mockResolvedValue([
      {
        types: ["text/html"],
        getType: vi.fn(async () => new Blob(["<p>web-html</p>"], { type: "text/html" })),
      },
    ]);
    const readText = vi.fn().mockResolvedValue("web-fallback");
    setMockNavigatorClipboard({ read, readText });

    const provider = await createClipboardProvider();
    const content = await provider.read();

    expect(content).toEqual({ text: "native-text", html: "<p>web-html</p>" });
    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("clipboard_read");
    expect(read).toHaveBeenCalledTimes(1);

    // We shouldn't need to fall back to plain text if rich read succeeded.
    expect(readText).not.toHaveBeenCalled();
  });

  it("read() decodes pngBase64 returned by clipboard_read into imagePng bytes", async () => {
    const pngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const pngBase64 = `data:image/png;base64,${Buffer.from(pngBytes).toString("base64")}`;

    const invoke = vi.fn().mockResolvedValue({
      text: "hello",
      html: "<p>hello</p>",
      pngBase64,
    });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const readText = vi.fn().mockResolvedValue("web-fallback");
    setMockNavigatorClipboard({ readText });

    const provider = await createClipboardProvider();
    const content = await provider.read();

    expect(content.text).toBe("hello");
    expect(content.html).toBe("<p>hello</p>");
    expect(content.imagePng).toEqual(pngBytes);
    expect((content as any).pngBase64).toBeUndefined();
    expect(readText).not.toHaveBeenCalled();
  });

  it("write() uses __TAURI__.core.invoke('clipboard_write', payload) when available", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    const legacyWriteText = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke }, clipboard: { writeText: legacyWriteText } };

    const webWriteText = vi.fn().mockResolvedValue(undefined);
    setMockNavigatorClipboard({ writeText: webWriteText });

    const provider = await createClipboardProvider();
    await provider.write({ text: "hello", html: "<p>hello</p>" });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_write");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      payload: { text: "hello", html: "<p>hello</p>" },
    });

    // No fallback should be needed.
    expect(legacyWriteText).not.toHaveBeenCalled();
    expect(webWriteText).not.toHaveBeenCalled();
  });

  it("write() encodes imagePng bytes as pngBase64 for clipboard_write payload", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    const legacyWriteText = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke }, clipboard: { writeText: legacyWriteText } };

    const webWriteText = vi.fn().mockResolvedValue(undefined);
    setMockNavigatorClipboard({ writeText: webWriteText });

    const pngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const expectedBase64 = Buffer.from(pngBytes).toString("base64");

    const provider = await createClipboardProvider();
    await provider.write({ text: "hello", imagePng: pngBytes });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_write");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      payload: { text: "hello", pngBase64: expectedBase64 },
    });

    expect(legacyWriteText).not.toHaveBeenCalled();
    expect(webWriteText).not.toHaveBeenCalled();
  });

  it("write() strips data: URL prefixes from payload.pngBase64 before sending to clipboard_write", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const pngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const rawBase64 = Buffer.from(pngBytes).toString("base64");

    const provider = await createClipboardProvider();
    await provider.write({ text: "hello", pngBase64: `data:image/png;base64,${rawBase64}` });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_write");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      payload: { text: "hello", pngBase64: rawBase64 },
    });
  });

  it("read() falls back to navigator.clipboard.readText when invoke is unavailable", async () => {
    (globalThis as any).__TAURI__ = {};

    const readText = vi.fn().mockResolvedValue("web-clipboard");
    setMockNavigatorClipboard({ readText });

    const provider = await createClipboardProvider();
    await expect(provider.read()).resolves.toEqual({ text: "web-clipboard" });

    expect(readText).toHaveBeenCalledTimes(1);
  });

  it("read() falls back to legacy __TAURI__.clipboard.readText when invoke is unavailable and web clipboard is unavailable", async () => {
    const legacyReadText = vi.fn().mockResolvedValue("legacy-clipboard");
    (globalThis as any).__TAURI__ = { clipboard: { readText: legacyReadText } };

    setMockNavigatorClipboard({});

    const provider = await createClipboardProvider();
    await expect(provider.read()).resolves.toEqual({ text: "legacy-clipboard" });

    expect(legacyReadText).toHaveBeenCalledTimes(1);
  });

  it("read() falls back to navigator.clipboard.readText when invoke rejects (e.g. unknown command)", async () => {
    const invoke = vi
      .fn()
      .mockRejectedValueOnce(new Error("unknown command: clipboard_read"))
      .mockRejectedValueOnce(new Error("unknown command: read_clipboard"));
    const legacyReadText = vi.fn().mockResolvedValue("legacy-clipboard");
    (globalThis as any).__TAURI__ = { core: { invoke }, clipboard: { readText: legacyReadText } };

    const readText = vi.fn().mockResolvedValue("web-clipboard");
    setMockNavigatorClipboard({ readText });

    const provider = await createClipboardProvider();
    await expect(provider.read()).resolves.toEqual({ text: "web-clipboard" });

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_read");
    expect(invoke.mock.calls[1]?.[0]).toBe("read_clipboard");

    expect(readText).toHaveBeenCalledTimes(1);
    expect(legacyReadText).not.toHaveBeenCalled();
  });

  it("read() falls back to legacy __TAURI__.clipboard.readText when invoke rejects and web clipboard is unavailable", async () => {
    const invoke = vi
      .fn()
      .mockRejectedValueOnce(new Error("unknown command: clipboard_read"))
      .mockRejectedValueOnce(new Error("unknown command: read_clipboard"));
    const legacyReadText = vi.fn().mockResolvedValue("legacy-clipboard");
    (globalThis as any).__TAURI__ = { core: { invoke }, clipboard: { readText: legacyReadText } };

    setMockNavigatorClipboard({});

    const provider = await createClipboardProvider();
    await expect(provider.read()).resolves.toEqual({ text: "legacy-clipboard" });

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_read");
    expect(invoke.mock.calls[1]?.[0]).toBe("read_clipboard");
    expect(legacyReadText).toHaveBeenCalledTimes(1);
  });

  it("write() falls back to legacy __TAURI__.clipboard.writeText when invoke rejects (e.g. unknown command)", async () => {
    const invoke = vi
      .fn()
      .mockRejectedValueOnce(new Error("unknown command: clipboard_write"))
      .mockRejectedValueOnce(new Error("unknown command: write_clipboard"));
    const legacyWriteText = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke }, clipboard: { writeText: legacyWriteText } };

    const webWriteText = vi.fn().mockResolvedValue(undefined);
    setMockNavigatorClipboard({ writeText: webWriteText });

    const provider = await createClipboardProvider();
    await expect(provider.write({ text: "fallback", html: "<p>fallback</p>" })).resolves.toBeUndefined();

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_write");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      payload: { text: "fallback", html: "<p>fallback</p>" },
    });
    expect(invoke.mock.calls[1]?.[0]).toBe("write_clipboard");
    expect(invoke.mock.calls[1]?.[1]).toMatchObject({ text: "fallback", html: "<p>fallback</p>" });

    expect(legacyWriteText).toHaveBeenCalledTimes(1);
    expect(legacyWriteText).toHaveBeenCalledWith("fallback");
    expect(webWriteText).not.toHaveBeenCalled();
  });

  it("write() falls back to legacy __TAURI__.clipboard.writeText when invoke is unavailable and web clipboard is unavailable", async () => {
    const legacyWriteText = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { clipboard: { writeText: legacyWriteText } };

    setMockNavigatorClipboard({});

    const provider = await createClipboardProvider();
    await expect(provider.write({ text: "fallback", html: "<p>fallback</p>" })).resolves.toBeUndefined();

    expect(legacyWriteText).toHaveBeenCalledTimes(1);
    expect(legacyWriteText).toHaveBeenCalledWith("fallback");
  });

  it("write() falls back to navigator.clipboard.writeText when invoke is unavailable", async () => {
    (globalThis as any).__TAURI__ = {};

    const webWriteText = vi.fn().mockResolvedValue(undefined);
    setMockNavigatorClipboard({ writeText: webWriteText });

    const provider = await createClipboardProvider();
    await expect(provider.write({ text: "fallback", html: "<p>fallback</p>" })).resolves.toBeUndefined();

    expect(webWriteText).toHaveBeenCalledTimes(1);
    expect(webWriteText).toHaveBeenCalledWith("fallback");
  });

  it("write() falls back to navigator.clipboard.writeText when invoke rejects and legacy clipboard API is unavailable", async () => {
    const invoke = vi
      .fn()
      .mockRejectedValueOnce(new Error("unknown command: clipboard_write"))
      .mockRejectedValueOnce(new Error("unknown command: write_clipboard"));
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const webWriteText = vi.fn().mockResolvedValue(undefined);
    setMockNavigatorClipboard({ writeText: webWriteText });

    const provider = await createClipboardProvider();
    await expect(provider.write({ text: "fallback", html: "<p>fallback</p>" })).resolves.toBeUndefined();

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke.mock.calls[0]?.[0]).toBe("clipboard_write");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      payload: { text: "fallback", html: "<p>fallback</p>" },
    });
    expect(invoke.mock.calls[1]?.[0]).toBe("write_clipboard");

    expect(webWriteText).toHaveBeenCalledTimes(1);
    expect(webWriteText).toHaveBeenCalledWith("fallback");
  });
});
