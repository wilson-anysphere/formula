export function computeQueryDependencies(query: any): any;

export class RefreshOrchestrator {
  constructor(options?: any);

  onEvent(listener: (event: any) => void): () => void;

  registerQuery(query: any): void;

  unregisterQuery(queryId: string): void;

  refreshAll(queryIds?: string[], reason?: any): {
    sessionId: string;
    queryIds: string[];
    promise: Promise<any>;
    cancel: () => void;
    cancelQuery?: (queryId: string) => void;
  };

  triggerOnOpen?(queryId?: string): any;
}
