import { JsonVectorStore } from "./jsonVectorStore.js";

/**
 * Node-only convenience wrapper around {@link JsonVectorStore} that persists to a file path.
 */
export class JsonFileVectorStore extends JsonVectorStore {
  constructor(opts: { filePath: string; dimension: number; autoSave?: boolean; resetOnCorrupt?: boolean });
  readonly filePath: string;
}
