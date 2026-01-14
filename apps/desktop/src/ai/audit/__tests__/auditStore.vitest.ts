/**
 * @vitest-environment node
 */

import { describe, expect, it, vi } from "vitest";

import type { AIAuditStore } from "@formula/ai-audit/browser";

vi.mock("sql.js/dist/sql-wasm.wasm?url", () => ({
  default: "/@fs/state/root/fake/sql-wasm.wasm",
}));

vi.mock("@formula/ai-audit/sqlite", () => ({
  SqliteAIAuditStore: {
    create: vi.fn(async (_opts: any): Promise<AIAuditStore> => ({
      logEntry: vi.fn(async () => {}),
      listEntries: vi.fn(async () => []),
    })),
  },
}));

import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
import { getDesktopAIAuditStore } from "../auditStore";

describe("getDesktopAIAuditStore", () => {
  it("coerces Vite /@fs wasm URLs to absolute file:// URLs in Node runtimes", async () => {
    const store = getDesktopAIAuditStore({ storageKey: "auditStoreTest:/@fs" });

    await store.logEntry({} as any);

    expect(SqliteAIAuditStore.create).toHaveBeenCalledTimes(1);
    const opts = (SqliteAIAuditStore.create as any).mock.calls[0][0];
    const locateFile = opts.locateFile as (file: string, prefix?: string) => string;

    // Use a non-existent filename so `import.meta.resolve` (if present) throws and we exercise
    // the `sqlWasmUrl` fallback (which we mock to a Vite `/@fs/...` URL).
    const wasmUrl = locateFile("missing.wasm");
    // In Node, sql.js uses `fs.readFileSync`, so we must return a filesystem path
    // (not a Vite dev-server URL or file:// URL).
    expect(wasmUrl).not.toMatch(/^file:\/\//);
    expect(wasmUrl).not.toContain("/@fs/");
    // Should be absolute (posix or Windows drive path).
    expect(wasmUrl).toMatch(/^\/|^[A-Za-z]:[\\/]/);
    expect(wasmUrl).toMatch(/sql-wasm\.wasm$/);
  });
});
