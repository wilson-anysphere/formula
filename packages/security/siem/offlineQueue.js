import { IndexedDbOfflineAuditQueue } from "./queue/indexeddb.js";

function isNodeRuntime() {
  return typeof process !== "undefined" && Boolean(process.versions?.node);
}

function isIndexedDbAvailable() {
  return typeof indexedDB !== "undefined";
}

async function loadNodeFsQueue() {
  const mod = await import(/* @vite-ignore */ "./queue/node_fs.js");
  return mod.NodeFsOfflineAuditQueue;
}

/**
 * Offline audit queue with crash-safe, resumable flushing.
 *
 * - Node: backed by segment files in `dirPath`
 * - Browser: backed by IndexedDB (`dbName`/`name`)
 */
export class OfflineAuditQueue {
  constructor(options = {}) {
    this.options = options;
    this.backend = options.backend;
    this.impl = null;
    this.implPromise = null;
  }

  async _getImpl() {
    if (this.impl) return this.impl;
    if (this.implPromise) return this.implPromise;

    const backend =
      this.backend ??
      (this.options.dirPath ? "fs" : isIndexedDbAvailable() ? "indexeddb" : isNodeRuntime() ? "fs" : null);

    if (!backend) {
      throw new Error("OfflineAuditQueue requires either dirPath (Node FS) or indexedDB availability (browser)");
    }

    if (backend === "indexeddb") {
      this.impl = new IndexedDbOfflineAuditQueue(this.options);
      return this.impl;
    }

    this.implPromise = (async () => {
      if (!this.options.dirPath) throw new Error("OfflineAuditQueue backend=fs requires dirPath");
      const NodeFsOfflineAuditQueue = await loadNodeFsQueue();
      return new NodeFsOfflineAuditQueue(this.options);
    })();

    this.impl = await this.implPromise;
    this.implPromise = null;
    return this.impl;
  }

  async enqueue(event) {
    const impl = await this._getImpl();
    return impl.enqueue(event);
  }

  async readAll() {
    const impl = await this._getImpl();
    return impl.readAll();
  }

  async clear() {
    const impl = await this._getImpl();
    return impl.clear();
  }

  async flushToExporter(exporter) {
    const impl = await this._getImpl();
    return impl.flushToExporter(exporter);
  }
}

