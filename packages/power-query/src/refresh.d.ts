export class RefreshManager {
  constructor(options?: any);

  onEvent(listener: (event: any) => void): () => void;

  registerQuery(query: any, policy?: any): void;

  unregisterQuery(queryId: string): void;

  refresh(queryId: string, reason?: any): { id: string; queryId: string; promise: Promise<any>; cancel: () => void };

  refreshAll(queryIds?: string[], reason?: any): { sessionId: string; promise: Promise<any>; cancel: () => void };

  triggerOnOpen(queryId?: string): void;

  dispose(): void;
}
