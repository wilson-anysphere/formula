import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import type { Logger } from "pino";
import type * as YTypes from "yjs";

import { Y } from "./yjs.js";

type PendingQueue = Promise<void>;

function encodeRecord(update: Uint8Array): Buffer {
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(update.byteLength, 0);
  return Buffer.concat([header, Buffer.from(update)]);
}

function decodeRecords(data: Buffer): Uint8Array[] {
  const out: Uint8Array[] = [];
  let offset = 0;
  while (offset + 4 <= data.length) {
    const len = data.readUInt32BE(offset);
    offset += 4;
    if (offset + len > data.length) break;
    out.push(new Uint8Array(data.subarray(offset, offset + len)));
    offset += len;
  }
  return out;
}

async function atomicWriteFile(filePath: string, contents: Buffer): Promise<void> {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmpPath, contents);
  try {
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "EEXIST" || code === "EPERM") {
      await fs.rm(filePath, { force: true });
      await fs.rename(tmpPath, filePath);
      return;
    }
    throw err;
  }
}

export class FilePersistence {
  private queues = new Map<string, PendingQueue>();
  private updateCounts = new Map<string, number>();
  private compactTimers = new Map<string, NodeJS.Timeout>();

  constructor(
    private readonly dir: string,
    private readonly logger: Logger,
    private readonly compactAfterUpdates: number
  ) {}

  private filePathForDoc(docName: string): string {
    const id = createHash("sha256").update(docName).digest("hex");
    return path.join(this.dir, `${id}.yjs`);
  }

  private enqueue(docName: string, task: () => Promise<void>): Promise<void> {
    const prev = this.queues.get(docName) ?? Promise.resolve();
    const next = prev
      .catch(() => {
        // Keep the queue alive even if a previous write failed.
      })
      .then(task);
    this.queues.set(docName, next);
    return next;
  }

  private scheduleCompaction(docName: string, doc: YTypes.Doc) {
    if (this.compactTimers.has(docName)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docName);
      void this.compactNow(docName, doc);
    }, 250);

    this.compactTimers.set(docName, timer);
  }

  private async compactNow(docName: string, doc: YTypes.Doc): Promise<void> {
    await this.enqueue(docName, async () => {
      const filePath = this.filePathForDoc(docName);
      const snapshot = Y.encodeStateAsUpdate(doc);
      await atomicWriteFile(filePath, encodeRecord(snapshot));
      this.updateCounts.set(docName, 0);
      this.logger.info({ docName }, "persistence_compacted");
    });
  }

  async bindState(docName: string, doc: YTypes.Doc): Promise<void> {
    await fs.mkdir(this.dir, { recursive: true });

    const filePath = this.filePathForDoc(docName);
    try {
      const data = await fs.readFile(filePath);
      for (const update of decodeRecords(data)) {
        Y.applyUpdate(doc, update);
      }
      this.logger.info({ docName }, "persistence_loaded");
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }

    const updateHandler = (update: Uint8Array) => {
      void this.enqueue(docName, async () => {
        await fs.appendFile(filePath, encodeRecord(update));
        const count = (this.updateCounts.get(docName) ?? 0) + 1;
        this.updateCounts.set(docName, count);

        if (count >= this.compactAfterUpdates) {
          this.scheduleCompaction(docName, doc);
        }
      });
    };

    doc.on("update", updateHandler);
    doc.on("destroy", () => {
      doc.off("update", updateHandler);
      const timer = this.compactTimers.get(docName);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docName);
    });
  }

  async writeState(docName: string, doc: YTypes.Doc): Promise<void> {
    await this.compactNow(docName, doc);
  }
}
