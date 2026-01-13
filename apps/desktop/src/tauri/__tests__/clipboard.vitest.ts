import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { readClipboard, writeClipboard } from "../clipboard";

describe("tauri/clipboard base64 normalization", () => {
  const base64 = Buffer.from([1, 2, 3]).toString("base64");
  let previousTauri: unknown;
  let hadTauri = false;

  beforeEach(() => {
    hadTauri = Object.prototype.hasOwnProperty.call(globalThis, "__TAURI__");
    previousTauri = (globalThis as any).__TAURI__;
  });

  afterEach(() => {
    if (hadTauri) (globalThis as any).__TAURI__ = previousTauri;
    else delete (globalThis as any).__TAURI__;
  });

  it("readClipboard strips data URI prefix + whitespace and decodes PNG bytes", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd !== "clipboard_read") throw new Error(`Unexpected invoke: ${cmd}`);
      return {
        pngBase64: `data:image/png;base64, \n ${base64} \n`,
      };
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const content = await readClipboard();
    expect(Array.from(content.imagePng ?? [])).toEqual([1, 2, 3]);
    expect(content.pngBase64).toBeUndefined();
  });

  it("readClipboard treats DATA: URL prefix case-insensitively", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd !== "clipboard_read") throw new Error(`Unexpected invoke: ${cmd}`);
      return {
        pngBase64: ` \n DATA:image/png;base64,${base64} \n`,
      };
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const content = await readClipboard();
    expect(Array.from(content.imagePng ?? [])).toEqual([1, 2, 3]);
    expect(content.pngBase64).toBeUndefined();
  });

  it("readClipboard ignores malformed data URL pngBase64 without a comma", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd !== "clipboard_read") throw new Error(`Unexpected invoke: ${cmd}`);
      return {
        pngBase64: "data:image/png;base64",
      };
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const content = await readClipboard();
    expect(content.imagePng).toBeUndefined();
    expect(content.pngBase64).toBeUndefined();
  });

  it("readClipboard falls back to legacy read_clipboard", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "clipboard_read") throw new Error("unsupported");
      if (cmd !== "read_clipboard") throw new Error(`Unexpected invoke: ${cmd}`);
      return {
        pngBase64: `data:image/png;base64,${base64}`,
      };
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const content = await readClipboard();
    expect(Array.from(content.imagePng ?? [])).toEqual([1, 2, 3]);
    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke).toHaveBeenNthCalledWith(1, "clipboard_read");
    expect(invoke).toHaveBeenNthCalledWith(2, "read_clipboard");
  });

  it("readClipboard preserves normalized pngBase64 when decoding fails", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd !== "clipboard_read") throw new Error(`Unexpected invoke: ${cmd}`);
      return { pngBase64: "data:image/png;base64, \n@@@@\n" };
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const content = await readClipboard();
    expect(content.imagePng).toBeUndefined();
    expect(content.pngBase64).toBe("@@@@");
  });

  it("writeClipboard normalizes legacy pngBase64 before invoking clipboard_write", async () => {
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      if (cmd !== "clipboard_write") throw new Error(`Unexpected invoke: ${cmd}`);
      expect(args).toEqual({ payload: { text: "", pngBase64: base64 } });
      return null;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    await writeClipboard({ text: "", pngBase64: `data:image/png;base64, \n ${base64} \n` });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("writeClipboard omits empty pngBase64 after normalization", async () => {
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      if (cmd !== "clipboard_write") throw new Error(`Unexpected invoke: ${cmd}`);
      expect(args).toEqual({ payload: { text: "hello" } });
      return null;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    await writeClipboard({ text: "hello", pngBase64: "data:image/png;base64," });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("writeClipboard treats DATA: URL prefix case-insensitively", async () => {
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      if (cmd !== "clipboard_write") throw new Error(`Unexpected invoke: ${cmd}`);
      expect(args).toEqual({ payload: { text: "", pngBase64: base64 } });
      return null;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    await writeClipboard({ text: "", pngBase64: ` \n DATA:image/png;base64,${base64} \n` });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("writeClipboard omits malformed data URL pngBase64 without a comma", async () => {
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      if (cmd !== "clipboard_write") throw new Error(`Unexpected invoke: ${cmd}`);
      expect(args).toEqual({ payload: { text: "hello" } });
      return null;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    await writeClipboard({ text: "hello", pngBase64: "data:image/png;base64" });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("writeClipboard falls back to legacy write_clipboard and normalizes base64", async () => {
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      if (cmd === "clipboard_write") throw new Error("unsupported");
      if (cmd !== "write_clipboard") throw new Error(`Unexpected invoke: ${cmd}`);
      expect(args).toEqual({
        text: "hello",
        html: undefined,
        rtf: undefined,
        image_png_base64: base64,
      });
      return null;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    await writeClipboard({ text: "hello", pngBase64: `data:image/png;base64,${base64}` });
    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke).toHaveBeenNthCalledWith(1, "clipboard_write", { payload: { text: "hello", pngBase64: base64 } });
    expect(invoke).toHaveBeenNthCalledWith(2, "write_clipboard", {
      text: "hello",
      html: undefined,
      rtf: undefined,
      image_png_base64: base64,
    });
  });
});
