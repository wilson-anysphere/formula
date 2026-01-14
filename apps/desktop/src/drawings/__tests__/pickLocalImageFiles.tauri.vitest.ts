import { afterEach, describe, expect, it, vi } from "vitest";

import { MAX_INSERT_IMAGE_BYTES } from "../insertImageLimits.js";
import { pickLocalImageFiles } from "../pickLocalImageFiles.js";

describe("pickLocalImageFiles (Tauri)", () => {
  afterEach(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = undefined;
  });

  it("uses __TAURI__.dialog.open and reads bytes via stat_file + read_binary_file", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["  /tmp/a.png  "]);
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      // Rust `stat_file` returns camelCase (`sizeBytes`); keep tests aligned with the real API shape.
      if (cmd === "stat_file") return { sizeBytes: 3 };
      if (cmd === "read_binary_file") {
        // eslint-disable-next-line no-undef
        return Buffer.from([1, 2, 3]).toString("base64");
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { dialog: { open }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: true });

    expect(open).toHaveBeenCalledTimes(1);
    expect((open.mock.calls[0]?.[0] as any)?.multiple).toBe(true);
    const filters = (open.mock.calls[0]?.[0] as any)?.filters ?? [];
    expect(filters[0]?.extensions).toEqual(["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"]);

    expect(calls.map((c) => c.cmd)).toEqual(["stat_file", "read_binary_file"]);
    expect(calls[0]?.args?.path).toBe("/tmp/a.png");
    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("a.png");
    expect(files[0]!.type).toBe("image/png");
    expect(files[0]!.size).toBe(3);
  });

  it("passes multiple=false to the Tauri dialog when requested", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => "/tmp/one.png");
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      // Rust `stat_file` returns camelCase (`sizeBytes`); keep tests aligned with the real API shape.
      if (cmd === "stat_file") return { sizeBytes: 1 };
      if (cmd === "read_binary_file") {
        // eslint-disable-next-line no-undef
        return Buffer.from([7]).toString("base64");
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { dialog: { open }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: false });

    expect(open).toHaveBeenCalledTimes(1);
    expect((open.mock.calls[0]?.[0] as any)?.multiple).toBe(false);
    expect(calls.map((c) => c.cmd)).toEqual(["stat_file", "read_binary_file"]);
    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("one.png");
    expect(files[0]!.type).toBe("image/png");
    expect(files[0]!.size).toBe(1);
  });

  it("uses read_binary_file_range for larger payloads", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["/tmp/big.jpg"]);
    const fileSize = 4 * 1024 * 1024 + 10; // > smallFileThreshold (4MiB)

    const base64Cache = new Map<number, string>();
    const base64Zeros = (length: number): string => {
      const len = Math.max(0, Math.trunc(length));
      const cached = base64Cache.get(len);
      if (cached) return cached;
      // eslint-disable-next-line no-undef
      const encoded = Buffer.from(new Uint8Array(len)).toString("base64");
      base64Cache.set(len, encoded);
      return encoded;
    };

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      // Rust `stat_file` returns camelCase (`sizeBytes`); keep tests aligned with the real API shape.
      if (cmd === "stat_file") return { sizeBytes: fileSize };
      if (cmd === "read_binary_file_range") {
        const length = Number(args?.length ?? 0);
        return base64Zeros(length);
      }
      if (cmd === "read_binary_file") {
        throw new Error("read_binary_file should not be used for large payloads");
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { dialog: { open }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: true });

    expect(open).toHaveBeenCalledTimes(1);
    expect(calls.some((c) => c.cmd === "read_binary_file_range")).toBe(true);
    expect(calls.some((c) => c.cmd === "read_binary_file")).toBe(false);

    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("big.jpg");
    expect(files[0]!.type).toBe("image/jpeg");
    expect(files[0]!.size).toBe(fileSize);
  });

  it("supports __TAURI__.plugins.dialog.open", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["/tmp/c.webp"]);
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      // Rust `stat_file` returns camelCase (`sizeBytes`); keep tests aligned with the real API shape.
      if (cmd === "stat_file") return { sizeBytes: 2 };
      if (cmd === "read_binary_file") {
        // eslint-disable-next-line no-undef
        return Buffer.from([9, 10]).toString("base64");
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { plugins: { dialog: { open } }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: true });

    expect(open).toHaveBeenCalledTimes(1);
    const filters = (open.mock.calls[0]?.[0] as any)?.filters ?? [];
    expect(filters[0]?.extensions).toEqual(["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"]);

    expect(calls.map((c) => c.cmd)).toEqual(["stat_file", "read_binary_file"]);
    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("c.webp");
    expect(files[0]!.type).toBe("image/webp");
    expect(files[0]!.size).toBe(2);
  });

  it("infers image/svg+xml for .svg files", async () => {
    const open = vi.fn(async () => ["/tmp/icon.svg"]);
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "stat_file") return { sizeBytes: 4 };
      if (cmd === "read_binary_file") {
        // eslint-disable-next-line no-undef
        return Buffer.from([0, 1, 2, 3]).toString("base64");
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { dialog: { open }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: true });
    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("icon.svg");
    expect(files[0]!.type).toBe("image/svg+xml");
    expect(files[0]!.size).toBe(4);
  });

  it("returns an oversized placeholder File when stat_file reports an oversized image", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["/tmp/huge.png"]);
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      // Rust `stat_file` returns camelCase (`sizeBytes`); keep tests aligned with the real API shape.
      if (cmd === "stat_file") return { sizeBytes: MAX_INSERT_IMAGE_BYTES + 1 };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { dialog: { open }, core: { invoke } };

    const files = await pickLocalImageFiles({ multiple: true });
    expect(open).toHaveBeenCalledTimes(1);
    expect(calls.map((c) => c.cmd)).toEqual(["stat_file"]);
    expect(calls[0]?.args).toEqual({ path: "/tmp/huge.png" });

    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("huge.png");
    expect(files[0]!.type).toBe("image/png");
    expect(files[0]!.size).toBeGreaterThan(MAX_INSERT_IMAGE_BYTES);
  });
});
