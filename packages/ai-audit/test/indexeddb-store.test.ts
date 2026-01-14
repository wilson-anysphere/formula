import { IDBKeyRange, indexedDB } from "fake-indexeddb";
import { describe, expect, it, afterEach, beforeEach } from "vitest";

import { IndexedDbAIAuditStore } from "../src/indexeddb-store.js";
import type { AIAuditEntry } from "../src/types.js";

import { randomUUID } from "node:crypto";

const originalGlobals = {
  indexedDB: (globalThis as any).indexedDB,
  IDBKeyRange: (globalThis as any).IDBKeyRange
};

beforeEach(() => {
  (globalThis as any).indexedDB = indexedDB;
  (globalThis as any).IDBKeyRange = IDBKeyRange;
});

afterEach(() => {
  if (originalGlobals.indexedDB === undefined) {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
  } else {
    (globalThis as any).indexedDB = originalGlobals.indexedDB;
  }

  if (originalGlobals.IDBKeyRange === undefined) {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;
  } else {
    (globalThis as any).IDBKeyRange = originalGlobals.IDBKeyRange;
  }
});

function createEntry(partial: Partial<AIAuditEntry> & Pick<AIAuditEntry, "id" | "timestamp_ms" | "session_id" | "mode">): AIAuditEntry {
  return {
    id: partial.id,
    timestamp_ms: partial.timestamp_ms,
    session_id: partial.session_id,
    workbook_id: partial.workbook_id,
    user_id: partial.user_id,
    mode: partial.mode,
    input: partial.input ?? { prompt: "hello" },
    model: partial.model ?? "unit-test-model",
    token_usage: partial.token_usage,
    latency_ms: partial.latency_ms,
    tool_calls: partial.tool_calls ?? [],
    verification: partial.verification,
    user_feedback: partial.user_feedback
  };
}

describe("IndexedDbAIAuditStore", () => {
  it("inserts and lists entries in newest-first order", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}` });
    const base = Date.now();

    await store.logEntry(
      createEntry({
        id: "older",
        timestamp_ms: base - 3000,
        session_id: "session-1",
        mode: "chat"
      })
    );

    await store.logEntry(
      createEntry({
        id: "newest",
        timestamp_ms: base - 1000,
        session_id: "session-1",
        mode: "chat"
      })
    );

    await store.logEntry(
      createEntry({
        id: "middle",
        timestamp_ms: base - 2000,
        session_id: "session-1",
        mode: "chat"
      })
    );

    const entries = await store.listEntries({ session_id: "session-1" });
    expect(entries.map((e) => e.id)).toEqual(["newest", "middle", "older"]);
  });

  it("falls back to JSON-safe cloning when entries contain uncloneable values (DataCloneError)", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}` });

    await store.logEntry(
      createEntry({
        id: "e1",
        timestamp_ms: Date.now(),
        session_id: "session-1",
        mode: "chat",
        input: { ok: true, fn: () => "nope" } as any,
      })
    );

    const entries = await store.listEntries({ session_id: "session-1" });
    expect(entries).toHaveLength(1);
    expect(entries[0]!.id).toBe("e1");
    expect(entries[0]!.input).toEqual({ ok: true });
  });

  it("filters by after_timestamp_ms (inclusive) and before_timestamp_ms (exclusive)", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}` });

    await store.logEntry(
      createEntry({
        id: "t1",
        timestamp_ms: 1000,
        session_id: "session-1",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "t2",
        timestamp_ms: 2000,
        session_id: "session-1",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "t3",
        timestamp_ms: 3000,
        session_id: "session-1",
        mode: "chat"
      })
    );

    expect((await store.listEntries({ after_timestamp_ms: 2000 })).map((e) => e.id)).toEqual(["t3", "t2"]);
    expect((await store.listEntries({ after_timestamp_ms: 3000 })).map((e) => e.id)).toEqual(["t3"]);

    expect((await store.listEntries({ before_timestamp_ms: 3000 })).map((e) => e.id)).toEqual(["t2", "t1"]);
    expect((await store.listEntries({ before_timestamp_ms: 2000 })).map((e) => e.id)).toEqual(["t1"]);

    expect((await store.listEntries({ after_timestamp_ms: 1500, before_timestamp_ms: 3000 })).map((e) => e.id)).toEqual([
      "t2"
    ]);
  });

  it("paginates stably with cursor across identical timestamps", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}` });

    await store.logEntry(
      createEntry({
        id: "a",
        timestamp_ms: 5000,
        session_id: "session-1",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "b",
        timestamp_ms: 5000,
        session_id: "session-1",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "c",
        timestamp_ms: 5000,
        session_id: "session-1",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "z",
        timestamp_ms: 4000,
        session_id: "session-1",
        mode: "chat"
      })
    );

    const page1 = await store.listEntries({ limit: 2 });
    expect(page1.map((e) => e.id)).toEqual(["c", "b"]);

    const page2 = await store.listEntries({
      limit: 2,
      cursor: { before_timestamp_ms: page1[1]!.timestamp_ms, before_id: page1[1]!.id }
    });
    expect(page2.map((e) => e.id)).toEqual(["a", "z"]);

    expect([...page1, ...page2].map((e) => e.id)).toEqual(["c", "b", "a", "z"]);
  });

  it("supports filtering by session_id/workbook_id/mode (single or array) and limit", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}` });
    const base = Date.now();

    await store.logEntry(
      createEntry({
        id: "e1",
        timestamp_ms: base - 4000,
        session_id: "session-1",
        workbook_id: "  workbook-1  ",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "e2",
        timestamp_ms: base - 3000,
        session_id: "session-1",
        workbook_id: "workbook-2",
        mode: "chat"
      })
    );
    await store.logEntry(
      createEntry({
        id: "e3",
        timestamp_ms: base - 2000,
        session_id: "session-2",
        workbook_id: "workbook-1",
        mode: "inline_edit"
      })
    );
    await store.logEntry(
      createEntry({
        id: "e4",
        timestamp_ms: base - 1000,
        session_id: "session-2",
        workbook_id: "workbook-1",
        mode: "chat"
      })
    );

    const session1 = await store.listEntries({ session_id: "session-1" });
    expect(session1.map((e) => e.id)).toEqual(["e2", "e1"]);

    const workbook1 = await store.listEntries({ workbook_id: "workbook-1" });
    expect(workbook1.map((e) => e.id)).toEqual(["e4", "e3", "e1"]);
    expect(workbook1.every((e) => e.workbook_id === "workbook-1")).toBe(true);

    const chatOnly = await store.listEntries({ mode: "chat" });
    expect(chatOnly.map((e) => e.id)).toEqual(["e4", "e2", "e1"]);

    const workbook1Modes = await store.listEntries({ workbook_id: "workbook-1", mode: ["chat", "inline_edit"] });
    expect(workbook1Modes.map((e) => e.id)).toEqual(["e4", "e3", "e1"]);

    const limited = await store.listEntries({ limit: 2 });
    expect(limited.map((e) => e.id)).toEqual(["e4", "e3"]);
  });

  it("enforces max_entries retention at write-time", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}`, max_entries: 2 });
    const base = Date.now();

    await store.logEntry(
      createEntry({
        id: "old",
        timestamp_ms: base - 3000,
        session_id: "session-1",
        workbook_id: "workbook-1",
        mode: "chat"
      })
    );

    await store.logEntry(
      createEntry({
        id: "middle",
        timestamp_ms: base - 2000,
        session_id: "session-2",
        workbook_id: "workbook-1",
        mode: "inline_edit"
      })
    );

    await store.logEntry(
      createEntry({
        id: "newest",
        timestamp_ms: base - 1000,
        session_id: "session-3",
        workbook_id: "workbook-2",
        mode: "chat"
      })
    );

    const all = await store.listEntries();
    expect(all.map((e) => e.id)).toEqual(["newest", "middle"]);
  });

  it("enforces max_age_ms retention at write-time", async () => {
    const store = new IndexedDbAIAuditStore({ db_name: `ai_audit_test_${randomUUID()}`, max_age_ms: 1000 });

    await store.logEntry(
      createEntry({
        id: "old",
        timestamp_ms: Date.now() - 10_000,
        session_id: "session-1",
        workbook_id: "workbook-1",
        mode: "chat",
        input: { prompt: "too old" }
      })
    );

    await store.logEntry(
      createEntry({
        id: "fresh",
        timestamp_ms: Date.now(),
        session_id: "session-1",
        workbook_id: "workbook-1",
        mode: "chat",
        input: { prompt: "fresh" }
      })
    );

    const entries = await store.listEntries();
    expect(entries.map((e) => e.id)).toEqual(["fresh"]);
  });
});
