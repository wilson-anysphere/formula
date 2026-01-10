import { appendFile, mkdir, readFile, unlink } from "node:fs/promises";
import path from "node:path";

export class OfflineAuditQueue {
  constructor(options) {
    if (!options || !options.dirPath) throw new Error("OfflineAuditQueue requires dirPath");
    this.dirPath = options.dirPath;
    this.filePath = options.filePath || path.join(this.dirPath, "audit-events.jsonl");
  }

  async ensureDir() {
    await mkdir(this.dirPath, { recursive: true });
  }

  async enqueue(event) {
    await this.ensureDir();
    const line = JSON.stringify(event) + "\n";
    await appendFile(this.filePath, line, "utf8");
  }

  async readAll() {
    try {
      const content = await readFile(this.filePath, "utf8");
      if (!content.trim()) return [];
      return content
        .split("\n")
        .filter(Boolean)
        .map((line) => JSON.parse(line));
    } catch (error) {
      if (error.code === "ENOENT") return [];
      throw error;
    }
  }

  async clear() {
    try {
      await unlink(this.filePath);
    } catch (error) {
      if (error.code === "ENOENT") return;
      throw error;
    }
  }

  async flushToExporter(exporter) {
    const events = await this.readAll();
    if (events.length === 0) return { sent: 0 };

    await exporter.sendBatch(events);
    await this.clear();
    return { sent: events.length };
  }
}
