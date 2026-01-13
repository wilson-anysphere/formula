import { beforeEach, describe, expect, it } from "vitest";

import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { migrateLocalStorageAuditEntriesToSqlite } from "../src/migrate.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";
import { InMemoryBinaryStorage } from "../src/storage.js";
import type { AIAuditEntry } from "../src/types.js";

function makeEntry(id: string, timestamp_ms: number): AIAuditEntry {
  return {
    id,
    timestamp_ms,
    session_id: "session-1",
    workbook_id: "workbook-1",
    mode: "chat",
    input: { prompt: `prompt-${id}` },
    model: "unit-test-model",
    tool_calls: []
  };
}

describe("migrateLocalStorageAuditEntriesToSqlite", () => {
  beforeEach(() => {
    globalThis.localStorage.clear();
  });

  it("migrates localStorage entries into sqlite and is idempotent", async () => {
    const key = `ai_audit_migrate_test_${Date.now()}`;
    const source = new LocalStorageAIAuditStore({ key });

    const entries = [
      makeEntry("entry-1", 1700000000000),
      makeEntry("entry-2", 1700000001000),
      makeEntry("entry-3", 1700000002000)
    ];
    for (const entry of entries) await source.logEntry(entry);

    const destination = await SqliteAIAuditStore.create({ storage: new InMemoryBinaryStorage() });

    await migrateLocalStorageAuditEntriesToSqlite({ source: { key }, destination });

    const migrated = await destination.listEntries();
    expect(migrated.map((e) => e.id)).toEqual(["entry-3", "entry-2", "entry-1"]);
    expect(new Map(migrated.map((e) => [e.id, e.timestamp_ms]))).toEqual(
      new Map(entries.map((e) => [e.id, e.timestamp_ms]))
    );

    // Re-running should not throw and should not duplicate.
    await expect(migrateLocalStorageAuditEntriesToSqlite({ source, destination })).resolves.toBeUndefined();
    const rerun = await destination.listEntries();
    expect(rerun.map((e) => e.id)).toEqual(["entry-3", "entry-2", "entry-1"]);
  });

  it("clears the source localStorage key when delete_source=true", async () => {
    const key = `ai_audit_migrate_delete_test_${Date.now()}`;
    const source = new LocalStorageAIAuditStore({ key });
    await source.logEntry(makeEntry("entry-delete-1", 1700000010000));

    expect(globalThis.localStorage.getItem(key)).not.toBeNull();

    const destination = await SqliteAIAuditStore.create({ storage: new InMemoryBinaryStorage() });
    await migrateLocalStorageAuditEntriesToSqlite({ source, destination, delete_source: true });

    expect(globalThis.localStorage.getItem(key)).toBeNull();
    const migrated = await destination.listEntries();
    expect(migrated.map((e) => e.id)).toEqual(["entry-delete-1"]);
  });
});

