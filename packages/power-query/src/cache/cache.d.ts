export class CacheManager {
  constructor(options?: any);

  prune(): Promise<void>;

  get(key: string, options?: any): Promise<unknown>;

  set(key: string, value: unknown, options?: any): Promise<void>;
}
