import { NodeFileBinaryStorage } from "./nodeFileBinaryStorage.js";
import { JsonVectorStore } from "./jsonVectorStore.js";

export class JsonFileVectorStore extends JsonVectorStore {
  /**
   * @param {{ filePath: string, dimension: number, autoSave?: boolean, resetOnCorrupt?: boolean }} opts
   */
  constructor(opts) {
    if (!opts?.filePath) throw new Error("JsonFileVectorStore requires filePath");
    const storage = new NodeFileBinaryStorage(opts.filePath);
    super({ storage, dimension: opts.dimension, autoSave: opts.autoSave, resetOnCorrupt: opts.resetOnCorrupt });
    this.filePath = opts.filePath;
  }
}
