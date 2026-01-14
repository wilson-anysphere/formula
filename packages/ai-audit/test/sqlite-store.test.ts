import { randomUUID } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import initSqlJs from "sql.js";
import { InMemoryBinaryStorage } from "../src/storage.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";
import { locateSqlJsFileNode } from "../src/sqlite-store.node.js";
import type { AIAuditEntry } from "../src/types.js";

describe("SqliteAIAuditStore", () => {
  const SQL_PROMISE = initSqlJs({ locateFile: locateSqlJsFileNode });

  async function createLegacyDatabaseBytes(setup: (db: any) => void): Promise<Uint8Array> {
    const SQL = await SQL_PROMISE;
    const db = new SQL.Database();
    setup(db);
    return db.export() as Uint8Array;
  }

  it("persists entries to sqlite via a binary storage adapter", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({ storage });

    const entry: AIAuditEntry = {
      id: randomUUID(),
      timestamp_ms: Date.now(),
      session_id: "session-1",
      mode: "chat",
      input: { prompt: "hello" },
      model: "unit-test-model",
      token_usage: { prompt_tokens: 10, completion_tokens: 5 },
      latency_ms: 123,
      tool_calls: [
        {
          name: "write_cell",
          parameters: { cell: "Sheet1!A1", value: 1 },
          requires_approval: true,
          approved: true,
          ok: true,
          duration_ms: 7
        }
      ],
      verification: {
        needs_tools: true,
        used_tools: true,
        verified: false,
        confidence: 0.25,
        warnings: ["No data tools were used; answer may be a guess."]
      },
      user_feedback: "accepted"
    };

    await store.logEntry(entry);

    const roundTrip = await SqliteAIAuditStore.create({ storage });
    const entries = await roundTrip.listEntries({ session_id: "session-1" });

    expect(entries.length).toBe(1);
    expect(entries[0]!.id).toBe(entry.id);
    expect(entries[0]!.model).toBe("unit-test-model");
    expect(entries[0]!.token_usage?.total_tokens).toBe(15);
    expect(entries[0]!.tool_calls[0]?.name).toBe("write_cell");
    expect(entries[0]!.tool_calls[0]?.approved).toBe(true);
    expect(entries[0]!.verification).toEqual(entry.verification);
    expect(entries[0]!.user_feedback).toBe("accepted");
  });

  it("serializes BigInt values as strings for JSON columns (input/tool_calls/verification)", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({ storage });

    const entry: AIAuditEntry = {
      id: randomUUID(),
      timestamp_ms: Date.now(),
      session_id: "session-bigint",
      mode: "chat",
      input: { big: 123n },
      model: "unit-test-model",
      tool_calls: [
        {
          name: "tool",
          parameters: { big: 456n },
        },
      ],
      verification: {
        needs_tools: false,
        used_tools: false,
        verified: true,
        confidence: 1,
        warnings: [],
        claims: [
          {
            claim: "bigint",
            verified: true,
            toolEvidence: { big: 789n },
          },
        ],
      },
    };

    await store.logEntry(entry);

    const roundTrip = await SqliteAIAuditStore.create({ storage });
    const entries = await roundTrip.listEntries({ session_id: "session-bigint" });

    expect(entries).toHaveLength(1);
    expect((entries[0]!.input as any).big).toBe("123");
    expect((entries[0]!.tool_calls[0]!.parameters as any).big).toBe("456");
    expect((entries[0]!.verification as any).claims[0].toolEvidence.big).toBe("789");
  });

  it("supports workbook/mode filtering and enforces retention at write-time", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({
      storage,
      retention: { max_entries: 2 }
    });

    const baseTime = Date.now();
    await store.logEntry({
      id: randomUUID(),
      timestamp_ms: baseTime - 3000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "older" },
      model: "unit-test-model",
      tool_calls: []
    });

    const keptInlineEditId = randomUUID();
    await store.logEntry({
      id: keptInlineEditId,
      timestamp_ms: baseTime - 2000,
      session_id: "session-2",
      workbook_id: "  workbook-1  ",
      mode: "inline_edit",
      input: { prompt: "mid" },
      model: "unit-test-model",
      tool_calls: []
    });

    const keptChatId = randomUUID();
    await store.logEntry({
      id: keptChatId,
      timestamp_ms: baseTime - 1000,
      session_id: "session-3",
      workbook_id: "workbook-2",
      mode: "chat",
      input: { prompt: "newest" },
      model: "unit-test-model",
      tool_calls: []
    });

    // max_entries=2 should have trimmed the oldest row.
    const all = await store.listEntries();
    expect(all.map((e) => e.id)).toEqual([keptChatId, keptInlineEditId]);

    const workbook1 = await store.listEntries({ workbook_id: "workbook-1" });
    expect(workbook1.map((e) => e.id)).toEqual([keptInlineEditId]);
    expect(workbook1[0]!.workbook_id).toBe("workbook-1");

    const chatOnly = await store.listEntries({ mode: "chat" });
    expect(chatOnly.map((e) => e.id)).toEqual([keptChatId]);

    const multiMode = await store.listEntries({ mode: ["chat", "inline_edit"] });
    expect(multiMode.map((e) => e.id)).toEqual([keptChatId, keptInlineEditId]);
  });

  it("supports max_age_ms retention", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({
      storage,
      retention: { max_age_ms: 1000 }
    });

    const oldId = randomUUID();
    await store.logEntry({
      id: oldId,
      timestamp_ms: Date.now() - 10_000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "too old" },
      model: "unit-test-model",
      tool_calls: []
    });

    const freshId = randomUUID();
    await store.logEntry({
      id: freshId,
      timestamp_ms: Date.now(),
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "fresh" },
      model: "unit-test-model",
      tool_calls: []
    });

    const entries = await store.listEntries();
    expect(entries.map((e) => e.id)).toEqual([freshId]);
  });

  it("backfills legacy workbook_id values from input_json on first open", async () => {
    const storage = new InMemoryBinaryStorage();
    const legacyId = randomUUID();

    await storage.save(
      await createLegacyDatabaseBytes((db) => {
        db.run(`
          CREATE TABLE ai_audit_log (
            id TEXT PRIMARY KEY,
            timestamp_ms INTEGER NOT NULL,
            session_id TEXT NOT NULL,
            user_id TEXT,
            mode TEXT NOT NULL,
            input_json TEXT NOT NULL,
            model TEXT NOT NULL,
            prompt_tokens INTEGER,
            completion_tokens INTEGER,
            total_tokens INTEGER,
            latency_ms INTEGER,
            tool_calls_json TEXT NOT NULL,
            user_feedback TEXT
          );
        `);

        db.run(
          `
            INSERT INTO ai_audit_log (
              id,
              timestamp_ms,
              session_id,
              user_id,
              mode,
              input_json,
              model,
              prompt_tokens,
              completion_tokens,
              total_tokens,
              latency_ms,
              tool_calls_json,
              user_feedback
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
          `,
          [
            legacyId,
            1700000000000,
            "session-legacy",
            null,
            "chat",
            JSON.stringify({ prompt: "hello", workbookId: "  workbook-legacy  " }),
            "unit-test-model",
            null,
            null,
            null,
            null,
            "[]",
            null
          ],
        );
      }),
    );

    const store = await SqliteAIAuditStore.create({ storage, locateFile: locateSqlJsFileNode });
    const entries = await store.listEntries({ workbook_id: "workbook-legacy" });

    expect(entries.map((e) => e.id)).toEqual([legacyId]);
    expect(entries[0]!.workbook_id).toBe("workbook-legacy");
  });

  it("backfills legacy workbook_id values from session_id when input_json is missing/malformed", async () => {
    const storage = new InMemoryBinaryStorage();
    const legacyId = randomUUID();
    const workbookId = "workbook-from-session";

    await storage.save(
      await createLegacyDatabaseBytes((db) => {
        db.run(`
          CREATE TABLE ai_audit_log (
            id TEXT PRIMARY KEY,
            timestamp_ms INTEGER NOT NULL,
            session_id TEXT NOT NULL,
            user_id TEXT,
            mode TEXT NOT NULL,
            input_json TEXT NOT NULL,
            model TEXT NOT NULL,
            prompt_tokens INTEGER,
            completion_tokens INTEGER,
            total_tokens INTEGER,
            latency_ms INTEGER,
            tool_calls_json TEXT NOT NULL,
            user_feedback TEXT
          );
        `);

        db.run(
          `
            INSERT INTO ai_audit_log (
              id,
              timestamp_ms,
              session_id,
              user_id,
              mode,
              input_json,
              model,
              prompt_tokens,
              completion_tokens,
              total_tokens,
              latency_ms,
              tool_calls_json,
              user_feedback
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
          `,
          [
            legacyId,
            1700000000000,
            `  ${workbookId}  :550e8400-e29b-41d4-a716-446655440000`,
            null,
            "chat",
            "{not-valid-json",
            "unit-test-model",
            null,
            null,
            null,
            null,
            "[]",
            null
          ],
        );
      }),
    );

    const store = await SqliteAIAuditStore.create({ storage, locateFile: locateSqlJsFileNode });
    const entries = await store.listEntries({ workbook_id: workbookId });

    expect(entries.map((e) => e.id)).toEqual([legacyId]);
    expect(entries[0]!.workbook_id).toBe(workbookId);
  });

  it("supports auto_persist=false by deferring storage writes until flush()", async () => {
    const storage = new InMemoryBinaryStorage();
    const saveSpy = vi.spyOn(storage, "save");
    const store = await SqliteAIAuditStore.create({ storage, auto_persist: false });

    const entry: AIAuditEntry = {
      id: randomUUID(),
      timestamp_ms: Date.now(),
      session_id: "session-buffered",
      mode: "chat",
      input: { prompt: "hello" },
      model: "unit-test-model",
      tool_calls: []
    };

    await store.logEntry(entry);

    expect(saveSpy).not.toHaveBeenCalled();

    await store.flush();
    expect(saveSpy).toHaveBeenCalledTimes(1);

    const roundTrip = await SqliteAIAuditStore.create({ storage });
    const entries = await roundTrip.listEntries({ session_id: "session-buffered" });
    expect(entries.map((e) => e.id)).toEqual([entry.id]);
  });

  it("debounces persistence when auto_persist_interval_ms is set", async () => {
    const storage = new InMemoryBinaryStorage();
    const saveSpy = vi.spyOn(storage, "save");
    const store = await SqliteAIAuditStore.create({ storage, auto_persist_interval_ms: 100 });

    vi.useFakeTimers();
    try {
      await store.logEntry({
        id: randomUUID(),
        timestamp_ms: Date.now(),
        session_id: "session-debounce",
        mode: "chat",
        input: { prompt: "one" },
        model: "unit-test-model",
        tool_calls: []
      });

      await store.logEntry({
        id: randomUUID(),
        timestamp_ms: Date.now(),
        session_id: "session-debounce",
        mode: "chat",
        input: { prompt: "two" },
        model: "unit-test-model",
        tool_calls: []
      });

      expect(saveSpy).toHaveBeenCalledTimes(0);

      await vi.advanceTimersByTimeAsync(99);
      expect(saveSpy).toHaveBeenCalledTimes(0);

      await vi.advanceTimersByTimeAsync(1);
      expect(saveSpy).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });
});
