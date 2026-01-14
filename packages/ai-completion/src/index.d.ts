export type SuggestionType = "formula" | "value" | "function_arg" | "range";

export interface Suggestion {
  text: string;
  displayText: string;
  type: SuggestionType;
  confidence: number;
  preview?: unknown;
}

export interface CellRefObject {
  row: number;
  col: number;
}

export type CellRef = CellRefObject | string;

export interface SurroundingCellsContext {
  getCellValue: (row: number, col: number, sheetName?: string) => unknown;
  getCacheKey?: () => string;
}

export interface CompletionContext {
  currentInput: string;
  cursorPosition: number;
  cellRef: CellRef;
  surroundingCells: SurroundingCellsContext;
}

export interface TabCompletionRequest {
  input: string;
  cursorPosition: number;
  cellA1: string;
  signal?: AbortSignal;
}

export interface TabCompletionClient {
  completeTabCompletion(req: TabCompletionRequest): Promise<string>;
}

export class CursorTabCompletionClient implements TabCompletionClient {
  constructor(options?: {
    baseUrl?: string;
    fetchImpl?: typeof fetch;
    timeoutMs?: number;
    getAuthHeaders?: () => Record<string, string> | Promise<Record<string, string>>;
  });
  completeTabCompletion(req: TabCompletionRequest): Promise<string>;
}

/**
 * Backwards-compatible alias for {@link CursorTabCompletionClient}.
 *
 * Prefer {@link CursorTabCompletionClient} for new code.
 */
export { CursorTabCompletionClient as CursorCompletionClient };

/**
 * Backwards-compatible alias for {@link TabCompletionClient}.
 *
 * Prefer {@link TabCompletionClient} for new code.
 */
export type CompletionClient = TabCompletionClient;

export class TabCompletionEngine {
  constructor(options?: {
    functionRegistry?: FunctionRegistry;
    parsePartialFormula?: ParsePartialFormula;
    completionClient?: TabCompletionClient | null;
    schemaProvider?: SchemaProvider | null;
    cache?: LRUCache<Suggestion[]>;
    cacheSize?: number;
    maxSuggestions?: number;
    completionTimeoutMs?: number;
    /**
     * Curated list of "starter" function stubs suggested when the user has only typed `=`.
     *
     * Callers may provide either a static list or a getter function (useful when starters depend on
     * runtime state like locale).
     */
    starterFunctions?: string[] | (() => string[]);
  });

  getSuggestions(
    context: CompletionContext,
    options?: { previewEvaluator?: PreviewEvaluator; signal?: AbortSignal }
  ): Promise<Suggestion[]>;
  buildCacheKey(context: CompletionContext): string;
}

export type ArgType = "range" | "value" | "number" | "string" | "boolean" | "any";

export interface FunctionArgSpec {
  name: string;
  type: ArgType;
  optional?: boolean;
  repeating?: boolean;
}

export interface FunctionSpec {
  name: string;
  description?: string;
  minArgs?: number;
  maxArgs?: number;
  args: FunctionArgSpec[];
  /**
   * Optional confidence boost applied when this function spec is suggested as a name completion.
   *
   * This is intended for host-specific registries that want to bias ordering (e.g. localized aliases).
   */
  completionBoost?: number;
}

export class FunctionRegistry {
  constructor(functions?: FunctionSpec[], options?: { catalog?: unknown });
  register(spec: FunctionSpec): void;
  list(): FunctionSpec[];
  getFunction(name: string): FunctionSpec | undefined;
  search(prefix: string, options?: { limit?: number }): FunctionSpec[];
  getArgType(functionName: string, argIndex: number): ArgType | undefined;
  isRangeArg(functionName: string, argIndex: number): boolean;
}

export interface PartialFormulaContext {
  isFormula: boolean;
  inFunctionCall: boolean;
  functionName?: string;
  argIndex?: number;
  expectingRange?: boolean;
  functionNamePrefix?: { text: string; start: number; end: number };
  currentArg?: { text: string; start: number; end: number };
}

/**
 * Signature accepted by {@link TabCompletionEngine}'s `parsePartialFormula` injection point.
 *
 * Callers may provide either a synchronous implementation (fast JS parser) or an async one
 * (e.g. locale-aware WASM partial parse).
 */
export type ParsePartialFormula = (
  input: string,
  cursorPosition: number,
  functionRegistry: { isRangeArg: (fnName: string, argIndex: number) => boolean }
) => PartialFormulaContext | Promise<PartialFormulaContext>;

export function parsePartialFormula(
  input: string,
  cursorPosition: number,
  functionRegistry: { isRangeArg: (fnName: string, argIndex: number) => boolean }
): PartialFormulaContext;

export class LRUCache<V = unknown> {
  constructor(maxEntries?: number);
  has(key: string): boolean;
  get(key: string): V | undefined;
  set(key: string, value: V): void;
  delete(key: string): boolean;
  clear(): void;
  readonly size: number;
}

export interface RangeSuggestion {
  range: string;
  confidence: number;
  reason: string;
}

export function suggestRanges(params: {
  currentArgText: string;
  cellRef: CellRef;
  surroundingCells: SurroundingCellsContext;
  sheetName?: string;
  maxScanRows?: number;
  maxScanCols?: number;
}): RangeSuggestion[];

export interface PatternSuggestion {
  text: string;
  confidence: number;
}

export function suggestPatternValues(params: {
  currentInput: string;
  cursorPosition: number;
  cellRef: CellRef;
  surroundingCells: SurroundingCellsContext;
  maxScanRows?: number;
  maxScanCols?: number;
}): PatternSuggestion[];

export function columnIndexToLetter(index: number): string;
export function columnLetterToIndex(letters: string): number;
export function normalizeCellRef(cellRef: CellRef): CellRefObject;
export function toA1(ref: CellRefObject): string;
export function parseA1(a1: string): CellRefObject | null;
export function isEmptyCell(value: unknown): boolean;

export interface NamedRangeInfo {
  name: string;
  range?: string;
}

export interface TableInfo {
  name: string;
  columns: string[];
  /**
   * Optional sheet identifier for the table.
   *
   * This is used by consumers (e.g. the desktop formula bar preview evaluator) to
   * resolve structured references into sheet-qualified A1 ranges.
   */
  sheetName?: string;
  /** 0-based inclusive bounds for the table range (if known). */
  startRow?: number;
  startCol?: number;
  endRow?: number;
  endCol?: number;
}

export interface SchemaProvider {
  getNamedRanges?: () => NamedRangeInfo[] | Promise<NamedRangeInfo[]>;
  getSheetNames?: () => string[] | Promise<string[]>;
  getTables?: () => TableInfo[] | Promise<TableInfo[]>;
  getCacheKey?: () => string;
}

export type PreviewEvaluator = (params: { suggestion: Suggestion; context: CompletionContext }) => unknown | Promise<unknown>;
