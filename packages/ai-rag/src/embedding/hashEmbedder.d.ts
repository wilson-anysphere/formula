export class HashEmbedder {
  constructor(opts?: { dimension?: number });
  readonly dimension: number;
  readonly name: string;
  embedTexts(texts: string[]): Promise<Float32Array[]>;
}

