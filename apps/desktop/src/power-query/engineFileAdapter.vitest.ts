import { afterEach, describe, expect, it, vi } from "vitest";

import { createDesktopQueryEngine } from "./engine.js";

describe("Power Query file adapter listDir", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("prefers backend list_dir when both core.invoke and the FS plugin are available", async () => {
    const invoke = vi.fn().mockResolvedValue([
      { path: "/tmp/folder/a.csv", name: "a.csv", size: 4, mtimeMs: 1000 },
      { path: "/tmp/folder/sub/b.json", name: "b.json", size: 2, mtimeMs: 1100 },
    ]);

    const readDir = vi.fn().mockResolvedValue([]);
    const readTextFile = vi.fn().mockResolvedValue("");
    const readFile = vi.fn().mockResolvedValue(new Uint8Array());

    (globalThis as any).__TAURI__ = {
      core: { invoke },
      plugin: {
        fs: {
          readDir,
          readTextFile,
          readFile,
        },
      },
    };

    const engine = createDesktopQueryEngine();
    const out = await engine.fileAdapter.listDir("/tmp/folder", { recursive: true });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("list_dir", { path: "/tmp/folder", recursive: true });
    expect(readDir).not.toHaveBeenCalled();
    expect(out).toHaveLength(2);
    expect(out.map((e) => e.path)).toEqual(["/tmp/folder/a.csv", "/tmp/folder/sub/b.json"]);
  });

  it("wraps list_dir resource limit errors with a user-friendly message", async () => {
    const invoke = vi.fn().mockRejectedValue(new Error("Directory listing exceeded limit (max 50000 entries)"));

    (globalThis as any).__TAURI__ = {
      core: { invoke },
      plugin: {
        fs: {
          readTextFile: vi.fn().mockResolvedValue(""),
          readFile: vi.fn().mockResolvedValue(new Uint8Array()),
          readDir: vi.fn().mockResolvedValue([]),
        },
      },
    };

    const engine = createDesktopQueryEngine();
    await expect(engine.fileAdapter.listDir("/tmp/folder", { recursive: true })).rejects.toThrow(
      /too many files|too deeply nested/i,
    );
  });

  it("falls back to __TAURI__.plugin.fs when a throwing __TAURI__.fs getter blocks direct FS access", async () => {
    const readTextFile = vi.fn().mockResolvedValue("hello");
    const readFile = vi.fn().mockResolvedValue(new Uint8Array());

    const tauri: any = {
      plugin: {
        fs: {
          readTextFile,
          readFile,
        },
      },
    };
    Object.defineProperty(tauri, "fs", {
      configurable: true,
      get() {
        throw new Error("Blocked fs access");
      },
    });
    (globalThis as any).__TAURI__ = tauri;

    const engine = createDesktopQueryEngine();
    await expect(engine.fileAdapter.readText("/tmp/file.txt")).resolves.toBe("hello");
    expect(readTextFile).toHaveBeenCalledTimes(1);
    expect(readTextFile).toHaveBeenCalledWith("/tmp/file.txt");
  });
});
