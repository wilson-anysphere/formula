import { redactAuditEvent } from "../redaction.js";

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function assertUuid(id) {
  if (typeof id !== "string" || !UUID_REGEX.test(id)) {
    throw new Error("audit event id must be a UUID");
  }
}

function redactIfConfigured(event, options) {
  if (options?.redact === false) return event;
  return redactAuditEvent(event, options?.redactionOptions);
}

const encoder = new TextEncoder();
function utf8ByteLength(text) {
  if (typeof Buffer !== "undefined") return Buffer.byteLength(text, "utf8");
  return encoder.encode(text).length;
}

function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function transactionDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
  });
}

async function runTransaction(db, storeNames, mode, fn) {
  const tx = db.transaction(storeNames, mode);
  try {
    const result = await fn(tx);
    await transactionDone(tx);
    return result;
  } catch (error) {
    try {
      tx.abort();
    } catch {
      // ignore
    }
    await transactionDone(tx).catch(() => {});
    throw error;
  }
}

function requireIndexedDb() {
  const idb = globalThis.indexedDB;
  if (!idb) throw new Error("indexedDB is not available in this environment");
  return idb;
}

async function openDatabase(dbName) {
  const indexedDb = requireIndexedDb();

  const request = indexedDb.open(dbName, 1);
  request.onupgradeneeded = () => {
    const db = request.result;
    if (!db.objectStoreNames.contains("events")) {
      const store = db.createObjectStore("events", { keyPath: "seq", autoIncrement: true });
      store.createIndex("status", "status", { unique: false });
    }
    if (!db.objectStoreNames.contains("meta")) {
      db.createObjectStore("meta", { keyPath: "key" });
    }
  };

  return requestToPromise(request);
}

async function getMetaNumber(tx, key, fallback = 0) {
  const store = tx.objectStore("meta");
  const record = await requestToPromise(store.get(key));
  const value = Number(record?.value);
  return Number.isFinite(value) ? value : fallback;
}

async function setMetaNumber(tx, key, value) {
  const store = tx.objectStore("meta");
  await requestToPromise(store.put({ key, value }));
}

export class IndexedDbOfflineAuditQueue {
  constructor(options = {}) {
    this.dbName = options.dbName ?? options.name ?? "siem-offline-audit-queue";
    this.maxBytes = options.maxBytes ?? 50 * 1024 * 1024;
    this.flushBatchSize = options.flushBatchSize ?? 250;
    this.redact = options.redact !== false;
    this.redactionOptions = options.redactionOptions;

    this._mutex = Promise.resolve();
    this.flushPromise = null;
    this.dbPromise = null;
  }

  _withMutex(fn) {
    const run = async () => fn();
    const result = this._mutex.then(run, run);
    this._mutex = result.then(
      () => undefined,
      () => undefined
    );
    return result;
  }

  async _getDb() {
    if (!this.dbPromise) this.dbPromise = openDatabase(this.dbName);
    return this.dbPromise;
  }

  async enqueue(event) {
    if (!event || typeof event !== "object") throw new Error("audit event must be an object");
    assertUuid(event.id);

    const safeEvent = redactIfConfigured(event, { redact: this.redact, redactionOptions: this.redactionOptions });
    const payload = JSON.stringify(safeEvent);
    const bytes = utf8ByteLength(payload) + 1;

    return this._withMutex(async () => {
      const db = await this._getDb();
      await runTransaction(db, ["events", "meta"], "readwrite", async (tx) => {
        const currentBytes = await getMetaNumber(tx, "bytes", 0);
        if (currentBytes + bytes > this.maxBytes) {
          const error = new Error("offline audit queue is full");
          error.code = "EQUEUEFULL";
          throw error;
        }

        const eventsStore = tx.objectStore("events");
        await requestToPromise(
          eventsStore.add({
            id: safeEvent.id,
            event: safeEvent,
            status: "pending",
            bytes,
            createdAtMs: Date.now(),
          })
        );
        await setMetaNumber(tx, "bytes", currentBytes + bytes);
      });
    });
  }

  async _reclaimInflight(tx) {
    const eventsStore = tx.objectStore("events");
    const index = eventsStore.index("status");
    const inflight = await requestToPromise(index.getAll("inflight"));
    for (const record of inflight) {
      await requestToPromise(eventsStore.put({ ...record, status: "pending" }));
    }
  }

  async _claimBatch() {
    const db = await this._getDb();
    return runTransaction(db, ["events"], "readwrite", async (tx) => {
      const eventsStore = tx.objectStore("events");
      const index = eventsStore.index("status");

      const pending = await requestToPromise(index.getAll("pending", this.flushBatchSize));
      for (const record of pending) {
        await requestToPromise(eventsStore.put({ ...record, status: "inflight" }));
      }
      return pending;
    });
  }

  async _releaseBatch(records) {
    if (records.length === 0) return;
    const db = await this._getDb();
    await runTransaction(db, ["events"], "readwrite", async (tx) => {
      const store = tx.objectStore("events");
      for (const record of records) {
        await requestToPromise(store.put({ ...record, status: "pending" }));
      }
    });
  }

  async _ackBatch(records) {
    if (records.length === 0) return;
    const db = await this._getDb();
    const bytesToRemove = records.reduce((sum, record) => sum + (Number(record.bytes) || 0), 0);

    await runTransaction(db, ["events", "meta"], "readwrite", async (tx) => {
      const store = tx.objectStore("events");
      for (const record of records) {
        await requestToPromise(store.delete(record.seq));
      }

      const currentBytes = await getMetaNumber(tx, "bytes", 0);
      await setMetaNumber(tx, "bytes", Math.max(0, currentBytes - bytesToRemove));
    });
  }

  async readAll() {
    const db = await this._getDb();
    return runTransaction(db, ["events"], "readonly", async (tx) => {
      const store = tx.objectStore("events");
      const all = await requestToPromise(store.getAll());
      return all
        .filter((record) => record.status === "pending" || record.status === "inflight")
        .sort((a, b) => a.seq - b.seq)
        .map((record) => record.event);
    });
  }

  async clear() {
    await this._withMutex(async () => {
      const db = await this._getDb();
      await runTransaction(db, ["events", "meta"], "readwrite", async (tx) => {
        await requestToPromise(tx.objectStore("events").clear());
        await requestToPromise(tx.objectStore("meta").clear());
        await setMetaNumber(tx, "bytes", 0);
      });
    });
  }

  async flushToExporter(exporter) {
    if (!exporter || typeof exporter.sendBatch !== "function") {
      throw new Error("flushToExporter requires exporter.sendBatch(events)");
    }

    if (this.flushPromise) return this.flushPromise;

    this.flushPromise = (async () => {
      const db = await this._getDb();
      await runTransaction(db, ["events"], "readwrite", (tx) => this._reclaimInflight(tx));

      let sent = 0;
      while (true) {
        const records = await this._claimBatch();
        if (records.length === 0) break;

        const events = records.map((record) => record.event);
        try {
          await exporter.sendBatch(events);
        } catch (error) {
          await this._releaseBatch(records);
          throw error;
        }

        await this._ackBatch(records);
        sent += events.length;
      }

      return { sent };
    })();

    try {
      return await this.flushPromise;
    } finally {
      this.flushPromise = null;
    }
  }
}

