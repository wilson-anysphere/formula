import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import type { Logger } from "pino";

export type TombstoneRecord = {
  deletedAtMs: number;
};

type TombstonesFileV1 = {
  schemaVersion: 1;
  tombstones: Record<string, TombstoneRecord>;
};

async function atomicWriteFile(filePath: string, contents: string): Promise<void> {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmpPath, contents, { encoding: "utf8", mode: 0o600 });
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

export function docKeyFromDocName(docName: string): string {
  return createHash("sha256").update(docName).digest("hex");
}

export class TombstoneStore {
  private readonly filePath: string;
  private initialized = false;
  private readonly tombstones = new Map<string, TombstoneRecord>();
  private writeQueue: Promise<void> = Promise.resolve();

  constructor(
    private readonly dataDir: string,
    private readonly logger: Logger
  ) {
    this.filePath = path.join(dataDir, "tombstones.json");
  }

  isInitialized(): boolean {
    return this.initialized;
  }

  async init(): Promise<void> {
    if (this.initialized) return;

    await fs.mkdir(this.dataDir, { recursive: true });

    try {
      const raw = await fs.readFile(this.filePath, "utf8");
      const parsed: unknown = JSON.parse(raw);
      if (
        !parsed ||
        typeof parsed !== "object" ||
        (parsed as { schemaVersion?: unknown }).schemaVersion !== 1 ||
        !(parsed as { tombstones?: unknown }).tombstones ||
        typeof (parsed as { tombstones?: unknown }).tombstones !== "object"
      ) {
        this.logger.warn({ filePath: this.filePath }, "tombstones_invalid_file");
      } else {
        const tombstones = (parsed as TombstonesFileV1).tombstones;
        for (const [docKey, record] of Object.entries(tombstones)) {
          if (
            typeof docKey === "string" &&
            record &&
            typeof record === "object" &&
            typeof (record as TombstoneRecord).deletedAtMs === "number"
          ) {
            this.tombstones.set(docKey, {
              deletedAtMs: (record as TombstoneRecord).deletedAtMs,
            });
          }
        }
      }
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }

    this.initialized = true;
  }

  count(): number {
    return this.tombstones.size;
  }

  has(docKey: string): boolean {
    return this.tombstones.has(docKey);
  }

  entries(): Array<[string, TombstoneRecord]> {
    return [...this.tombstones.entries()];
  }

  async set(docKey: string, deletedAtMs: number = Date.now()): Promise<void> {
    this.tombstones.set(docKey, { deletedAtMs });
    await this.persist();
  }

  async delete(docKey: string): Promise<void> {
    if (!this.tombstones.delete(docKey)) return;
    await this.persist();
  }

  async sweepExpired(ttlMs: number, nowMs: number = Date.now()): Promise<{
    expiredDocKeys: string[];
  }> {
    const expiredDocKeys: string[] = [];
    if (ttlMs <= 0) return { expiredDocKeys };

    const cutoff = nowMs - ttlMs;
    for (const [docKey, record] of this.tombstones.entries()) {
      if (record.deletedAtMs <= cutoff) {
        this.tombstones.delete(docKey);
        expiredDocKeys.push(docKey);
      }
    }

    if (expiredDocKeys.length > 0) await this.persist();
    return { expiredDocKeys };
  }

  private async persist(): Promise<void> {
    this.writeQueue = this.writeQueue
      .catch(() => {
        // keep queue alive
      })
      .then(() => this.persistNow());
    return this.writeQueue;
  }

  private async persistNow(): Promise<void> {
    await fs.mkdir(this.dataDir, { recursive: true });

    const sorted = [...this.tombstones.entries()].sort(([a], [b]) =>
      a.localeCompare(b)
    );
    const json: TombstonesFileV1 = {
      schemaVersion: 1,
      tombstones: Object.fromEntries(sorted),
    };
    await atomicWriteFile(this.filePath, `${JSON.stringify(json)}\n`);
  }
}
