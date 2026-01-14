import { describe, expect, it, vi } from "vitest";

import { installOpenFileIpc } from "../openFileIpc";

describe("openFileIpc", () => {
  it("emits open-file-ready only after the open-file listener is registered", async () => {
    let resolveListen: ((unlisten: () => void) => void) | null = null;
    const listen = vi.fn(() => {
      return new Promise<() => void>((resolve) => {
        resolveListen = resolve;
      });
    });

    const emit = vi.fn();
    const onOpenPath = vi.fn();

    installOpenFileIpc({ listen, emit, onOpenPath });

    expect(listen).toHaveBeenCalledTimes(1);
    expect(listen).toHaveBeenCalledWith("open-file", expect.any(Function));
    expect(emit).not.toHaveBeenCalled();

    resolveListen?.(() => {});
    await Promise.resolve();

    expect(emit).toHaveBeenCalledTimes(1);
    expect(emit).toHaveBeenCalledWith("open-file-ready");
  });

  it("dispatches each non-empty string path in the open-file payload", () => {
    let handler: ((event: any) => void) | null = null;
    const listen = vi.fn((_: string, fn: (event: any) => void) => {
      handler = fn;
      return Promise.resolve(() => {});
    });

    const onOpenPath = vi.fn();

    installOpenFileIpc({ listen, emit: null, onOpenPath });

    handler?.({ payload: ["book.xlsx", "  data.csv  ", " ", 123] });

    expect(onOpenPath).toHaveBeenCalledTimes(2);
    expect(onOpenPath).toHaveBeenNthCalledWith(1, "book.xlsx");
    expect(onOpenPath).toHaveBeenNthCalledWith(2, "data.csv");
  });
});
