export class OpenAIEmbedder {
  constructor(opts: { apiKey: string; model: string; baseUrl?: string });
  readonly name: string;
  readonly dimension: number | null;
  embedTexts(texts: string[]): Promise<number[][]>;
}

