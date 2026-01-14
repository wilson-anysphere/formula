import { afterEach, describe, expect, it, vi } from "vitest";

import { pickLocalImageFiles } from "../pickLocalImageFiles.js";

describe("pickLocalImageFiles (Tauri)", () => {
  afterEach(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = undefined;
  });

  it("uses __TAURI__.dialog.open and reads bytes via stat_file + read_binary_file", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["/tmp/a.png"]);
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "stat_file") return { size_bytes: 3 };
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
    const filters = (open.mock.calls[0]?.[0] as any)?.filters ?? [];
    expect(filters[0]?.extensions).toEqual(["png", "jpg", "jpeg", "gif", "bmp", "webp"]);

    expect(calls.map((c) => c.cmd)).toEqual(["stat_file", "read_binary_file"]);
    expect(files).toHaveLength(1);
    expect(files[0]!.name).toBe("a.png");
    expect(files[0]!.type).toBe("image/png");
    expect(files[0]!.size).toBe(3);
  });

  it("uses read_binary_file_range for larger payloads", async () => {
    const calls: Array<{ cmd: string; args: any }> = [];

    const open = vi.fn(async () => ["/tmp/big.jpg"]);
    const fileSize = 4 * 1024 * 1024 + 10; // > smallFileThreshold (4MiB)

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "stat_file") return { size_bytes: fileSize };
      if (cmd === "read_binary_file_range") {
        const length = Number(args?.length ?? 0);
        return new Uint8Array(length);
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
});

