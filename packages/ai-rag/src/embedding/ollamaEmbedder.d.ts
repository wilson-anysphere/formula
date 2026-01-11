export class OllamaEmbedder {
  constructor(opts: { model: string; host?: string; dimension?: number });
  readonly name: string;
  /**
   * Dimension depends on model; populated after first embedding call when omitted.
   */
  readonly dimension: number | null;
  embedTexts(texts: string[]): Promise<number[][]>;
}

