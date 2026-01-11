import type { Query } from "./model.js";
import type { ArrowTableAdapter } from "./arrowTable.js";
import type { DataTable } from "./table.js";

export type QueryExecutionContext = Record<string, any>;

export type QueryExecutionResult = { table: any; meta: any };

export class QueryEngine {
  constructor(options?: any);

  createSession(options?: { now?: () => number }): any;

  executeQuery(query: Query, context?: QueryExecutionContext, options?: any): Promise<DataTable | ArrowTableAdapter>;

  executeQueryWithMeta(query: Query, context?: QueryExecutionContext, options?: any): Promise<QueryExecutionResult>;

  executeQueryWithMetaInSession(
    query: Query,
    context: QueryExecutionContext,
    options: any,
    session: any,
  ): Promise<QueryExecutionResult>;

  executeQueryStreaming(query: Query, context: QueryExecutionContext, options: any): Promise<any>;

  getCacheKey(query: Query, context?: QueryExecutionContext, options?: any): Promise<string>;

  invalidateQueryCache(query: Query, context?: QueryExecutionContext, options?: any): Promise<void>;
}

export function parseCsv(text: string, options?: any): any;
export function parseCsvCell(text: string, options?: any): any;
