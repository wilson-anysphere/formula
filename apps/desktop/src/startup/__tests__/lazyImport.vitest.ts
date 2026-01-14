import { describe, expect, it, vi } from "vitest";

import { createLazyImport } from "../lazyImport.js";

describe("createLazyImport", () => {
  it("caches the module promise across calls", async () => {
    const importer = vi.fn(async () => ({ value: 42 }));
    const load = createLazyImport(importer, { label: "test" });

    const [a, b] = await Promise.all([load(), load()]);
    expect(a).toEqual({ value: 42 });
    expect(b).toEqual({ value: 42 });
    expect(importer).toHaveBeenCalledTimes(1);
  });

  it("invokes onError and allows retry after failure", async () => {
    const errors: unknown[] = [];
    let attempt = 0;
    const importer = vi.fn(async () => {
      attempt += 1;
      if (attempt === 1) throw new Error("boom");
      return { ok: true };
    });

    const load = createLazyImport(importer, {
      label: "test",
      onError: (err) => errors.push(err),
    });

    const first = await load();
    expect(first).toBeNull();
    expect(errors).toHaveLength(1);
    expect(importer).toHaveBeenCalledTimes(1);

    const second = await load();
    expect(second).toEqual({ ok: true });
    expect(importer).toHaveBeenCalledTimes(2);
  });
});

