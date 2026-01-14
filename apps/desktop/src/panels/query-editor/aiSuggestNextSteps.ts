import type {
  Aggregation,
  AndPredicate,
  ArrowTableAdapter,
  ComparisonPredicate,
  DataTable,
  FilterPredicate,
  NotPredicate,
  OrPredicate,
  Query,
  QueryOperation,
  SortSpec,
} from "@formula/power-query";
import { stableStringify } from "@formula/power-query";

import { getDesktopLLMClient, getDesktopModel } from "../../ai/llm/desktopLLMClient.js";

type PreviewTable = DataTable | ArrowTableAdapter | null;

const SYSTEM_PROMPT = `You are an expert at Power Query transformations.

Given a user's intent, suggest 1-3 next query operations.

Return ONLY valid JSON (no markdown, no code fences): a JSON array of QueryOperation objects.

Allowed operation types and shapes:
- filterRows: { "type": "filterRows", "predicate": FilterPredicate }
- sortRows: { "type": "sortRows", "sortBy": Array<{ "column": string, "direction"?: "ascending"|"descending", "nulls"?: "first"|"last" }> }
- removeColumns: { "type": "removeColumns", "columns": string[] }
- selectColumns: { "type": "selectColumns", "columns": string[] }
- renameColumn: { "type": "renameColumn", "oldName": string, "newName": string }
- changeType: { "type": "changeType", "column": string, "newType": "any"|"string"|"number"|"boolean"|"date" }
- splitColumn: { "type": "splitColumn", "column": string, "delimiter": string }
- groupBy: { "type": "groupBy", "groupColumns": string[], "aggregations": Array<{ "column": string, "op": "sum"|"count"|"average"|"min"|"max"|"countDistinct", "as"?: string }> }
- addColumn: { "type": "addColumn", "name": string, "formula": string }
- take: { "type": "take", "count": number }
- distinctRows: { "type": "distinctRows", "columns": string[] | null }
- removeRowsWithErrors: { "type": "removeRowsWithErrors", "columns": string[] | null }
- fillDown: { "type": "fillDown", "columns": string[] }
- replaceValues: { "type": "replaceValues", "column": string, "find": any, "replace": any }

FilterPredicate shapes:
- comparison: { "type": "comparison", "column": string, "operator": "equals"|"notEquals"|"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"contains"|"startsWith"|"endsWith"|"isNull"|"isNotNull", "value"?: any, "caseSensitive"?: boolean }
- and: { "type": "and", "predicates": FilterPredicate[] }
- or: { "type": "or", "predicates": FilterPredicate[] }
- not: { "type": "not", "predicate": FilterPredicate }

Rules:
- Use only column names from the provided schema.
- If the schema is empty, suggest only schema-independent operations (e.g. take, distinctRows with columns=null, removeRowsWithErrors with columns=null).
- Keep operations minimally valid (required fields present, correct types).
- Do not invent column names that do not exist.`;

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function stripCodeFences(text: string): string {
  const trimmed = text.trim();
  const match = /^```(?:json)?\s*([\s\S]*?)\s*```$/i.exec(trimmed);
  return match ? match[1]!.trim() : trimmed;
}

function extractJsonCandidate(text: string): string {
  const cleaned = stripCodeFences(text);
  const startArray = cleaned.indexOf("[");
  const endArray = cleaned.lastIndexOf("]");
  if (startArray >= 0 && endArray >= startArray) {
    return cleaned.slice(startArray, endArray + 1);
  }
  const startObj = cleaned.indexOf("{");
  const endObj = cleaned.lastIndexOf("}");
  if (startObj >= 0 && endObj >= startObj) {
    return cleaned.slice(startObj, endObj + 1);
  }
  return cleaned;
}

function parseJson(text: string): unknown {
  const candidate = extractJsonCandidate(text);
  return JSON.parse(candidate);
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((v) => typeof v === "string");
}

function isDataType(value: unknown): value is "any" | "string" | "number" | "boolean" | "date" {
  return value === "any" || value === "string" || value === "number" || value === "boolean" || value === "date";
}

function isComparisonOperator(value: unknown): value is ComparisonPredicate["operator"] {
  return (
    value === "equals" ||
    value === "notEquals" ||
    value === "greaterThan" ||
    value === "greaterThanOrEqual" ||
    value === "lessThan" ||
    value === "lessThanOrEqual" ||
    value === "contains" ||
    value === "startsWith" ||
    value === "endsWith" ||
    value === "isNull" ||
    value === "isNotNull"
  );
}

function validateColumn(name: unknown, allowed: Set<string>): name is string {
  if (typeof name !== "string") return false;
  if (!name.trim()) return false;
  return allowed.has(name);
}

function isAggregationOp(value: unknown): value is Aggregation["op"] {
  return value === "sum" || value === "count" || value === "average" || value === "min" || value === "max" || value === "countDistinct";
}

function coerceFilterPredicate(value: unknown, allowedColumns: Set<string>): FilterPredicate | null {
  if (!isPlainObject(value)) return null;
  const type = value.type;
  if (type === "comparison") {
    if (!validateColumn(value.column, allowedColumns)) return null;
    if (!isComparisonOperator(value.operator)) return null;
    const predicate: ComparisonPredicate = {
      type: "comparison",
      column: value.column,
      operator: value.operator,
    };
    const maybeValue = (value as { value?: unknown }).value;
    if ("value" in value) predicate.value = maybeValue;
    const maybeCaseSensitive = (value as { caseSensitive?: unknown }).caseSensitive;
    if (typeof maybeCaseSensitive === "boolean") predicate.caseSensitive = maybeCaseSensitive;
    return predicate;
  }
  if (type === "and" || type === "or") {
    const predicates = (value as { predicates?: unknown }).predicates;
    if (!Array.isArray(predicates)) return null;
    const out = predicates
      .map((p: unknown) => coerceFilterPredicate(p, allowedColumns))
      .filter((p): p is FilterPredicate => Boolean(p));
    if (out.length === 0) return null;
    return type === "and" ? ({ type: "and", predicates: out } satisfies AndPredicate) : ({ type: "or", predicates: out } satisfies OrPredicate);
  }
  if (type === "not") {
    const inner = coerceFilterPredicate((value as { predicate?: unknown }).predicate, allowedColumns);
    if (!inner) return null;
    return { type: "not", predicate: inner } satisfies NotPredicate;
  }
  return null;
}

function coerceQueryOperation(value: unknown, allowedColumns: Set<string>): QueryOperation | null {
  if (!isPlainObject(value)) return null;
  const type = value.type;
  if (typeof type !== "string") return null;

  // When we don't have a preview schema, avoid suggesting operations that
  // reference column names or depend on column layout. Even if an operation could
  // be technically valid (e.g. `removeColumns: []`), it's not helpful and may
  // confuse users because the UI intentionally disables schema-dependent actions
  // in that state.
  if (allowedColumns.size === 0) {
    if (type === "take") {
      const count = (value as { count?: unknown }).count;
      if (typeof count !== "number" || !Number.isFinite(count)) return null;
      return { type: "take", count };
    }
    if (type === "distinctRows") {
      const columns = (value as { columns?: unknown }).columns;
      if (columns == null) return { type: "distinctRows", columns: null };
      return null;
    }
    if (type === "removeRowsWithErrors") {
      const columns = (value as { columns?: unknown }).columns;
      if (columns == null) return { type: "removeRowsWithErrors", columns: null };
      return null;
    }
    return null;
  }

  switch (type) {
    case "take": {
      const count = (value as { count?: unknown }).count;
      if (typeof count !== "number" || !Number.isFinite(count)) return null;
      return { type: "take", count };
    }
    case "distinctRows": {
      const columns = (value as { columns?: unknown }).columns;
      if (columns == null) return { type: "distinctRows", columns: null };
      if (!isStringArray(columns)) return null;
      if (columns.some((c) => !allowedColumns.has(c))) return null;
      return { type: "distinctRows", columns };
    }
    case "removeRowsWithErrors": {
      const columns = (value as { columns?: unknown }).columns;
      if (columns == null) return { type: "removeRowsWithErrors", columns: null };
      if (!isStringArray(columns)) return null;
      if (columns.some((c) => !allowedColumns.has(c))) return null;
      return { type: "removeRowsWithErrors", columns };
    }
    case "removeColumns": {
      const columns = (value as { columns?: unknown }).columns;
      if (!isStringArray(columns)) return null;
      if (columns.some((c) => !allowedColumns.has(c))) return null;
      return { type: "removeColumns", columns };
    }
    case "selectColumns": {
      const columns = (value as { columns?: unknown }).columns;
      if (!isStringArray(columns)) return null;
      if (columns.some((c) => !allowedColumns.has(c))) return null;
      return { type: "selectColumns", columns };
    }
    case "renameColumn": {
      const oldName = (value as { oldName?: unknown }).oldName;
      const newName = (value as { newName?: unknown }).newName;
      if (!validateColumn(oldName, allowedColumns)) return null;
      if (typeof newName !== "string" || !newName.trim()) return null;
      return { type: "renameColumn", oldName, newName };
    }
    case "changeType": {
      const column = (value as { column?: unknown }).column;
      const newType = (value as { newType?: unknown }).newType;
      if (!validateColumn(column, allowedColumns)) return null;
      if (!isDataType(newType)) return null;
      return { type: "changeType", column, newType };
    }
    case "splitColumn": {
      const column = (value as { column?: unknown }).column;
      const delimiter = (value as { delimiter?: unknown }).delimiter;
      if (!validateColumn(column, allowedColumns)) return null;
      if (typeof delimiter !== "string") return null;
      return { type: "splitColumn", column, delimiter };
    }
    case "fillDown": {
      const columns = (value as { columns?: unknown }).columns;
      if (!isStringArray(columns)) return null;
      if (columns.some((c) => !allowedColumns.has(c))) return null;
      return { type: "fillDown", columns };
    }
    case "replaceValues": {
      const column = (value as { column?: unknown }).column;
      if (!validateColumn(column, allowedColumns)) return null;
      if (!("find" in value) || !("replace" in value)) return null;
      const record = value as { find?: unknown; replace?: unknown };
      return { type: "replaceValues", column, find: record.find, replace: record.replace };
    }
    case "addColumn": {
      const name = (value as { name?: unknown }).name;
      const formula = (value as { formula?: unknown }).formula;
      if (typeof name !== "string" || !name.trim()) return null;
      if (typeof formula !== "string" || !formula.trim()) return null;
      return { type: "addColumn", name, formula };
    }
    case "sortRows": {
      const sortBy = (value as { sortBy?: unknown }).sortBy;
      if (!Array.isArray(sortBy) || sortBy.length === 0) return null;
      const out: SortSpec[] = [];
      for (const spec of sortBy) {
        if (!isPlainObject(spec)) return null;
        const column = (spec as { column?: unknown }).column;
        if (!validateColumn(column, allowedColumns)) return null;
        const direction = (spec as { direction?: unknown }).direction;
        const nulls = (spec as { nulls?: unknown }).nulls;
        if (direction != null && direction !== "ascending" && direction !== "descending") return null;
        if (nulls != null && nulls !== "first" && nulls !== "last") return null;
        out.push({
          column,
          ...(direction != null ? { direction } : {}),
          ...(nulls != null ? { nulls } : {}),
        });
      }
      return { type: "sortRows", sortBy: out };
    }
    case "filterRows": {
      const predicate = coerceFilterPredicate((value as { predicate?: unknown }).predicate, allowedColumns);
      if (!predicate) return null;
      return { type: "filterRows", predicate };
    }
    case "groupBy": {
      const groupColumns = (value as { groupColumns?: unknown }).groupColumns;
      const aggregations = (value as { aggregations?: unknown }).aggregations;
      if (!isStringArray(groupColumns) || groupColumns.length === 0) return null;
      if (groupColumns.some((c) => !allowedColumns.has(c))) return null;
      if (!Array.isArray(aggregations) || aggregations.length === 0) return null;
      const outAggs: Aggregation[] = [];
      for (const agg of aggregations) {
        if (!isPlainObject(agg)) return null;
        const column = (agg as { column?: unknown }).column;
        const op = (agg as { op?: unknown }).op;
        const as = (agg as { as?: unknown }).as;
        if (!validateColumn(column, allowedColumns)) return null;
        if (!isAggregationOp(op)) return null;
        if (as != null && (typeof as !== "string" || !as.trim())) return null;
        outAggs.push(as ? { column, op, as } : { column, op });
      }
      return { type: "groupBy", groupColumns, aggregations: outAggs };
    }
    default:
      return null;
  }
}

function coerceOperations(value: unknown, allowedColumns: Set<string>): QueryOperation[] {
  const operations: unknown[] = Array.isArray(value)
    ? value
    : isPlainObject(value) && Array.isArray((value as { operations?: unknown }).operations)
      ? ((value as { operations?: unknown }).operations as unknown[])
      : [];
  return operations
    .map((op) => coerceQueryOperation(op, allowedColumns))
    .filter((op): op is QueryOperation => Boolean(op))
    .slice(0, 8);
}

function buildUserPrompt(args: { intent: string; query: Query; preview: PreviewTable }): string {
  const schema = args.preview?.columns?.map((c) => ({ name: c.name, type: c.type })) ?? [];
  const previewInfo = args.preview ? { rowCount: args.preview.rowCount, columnCount: args.preview.columnCount } : null;
  const steps = args.query.steps.map((s) => s.operation);

  return `User intent: ${args.intent}

Current steps (operations only):
${stableStringify(steps)}

Preview schema:
${stableStringify(schema)}

Preview info:
${previewInfo ? stableStringify(previewInfo) : "null"}`;
}

export async function suggestQueryNextSteps(intent: string, ctx: { query: Query; preview: PreviewTable }): Promise<QueryOperation[]> {
  const client = getDesktopLLMClient();
  const model = getDesktopModel();

  const request = {
    model,
    temperature: 0.2,
    maxTokens: 500,
    messages: [
      { role: "system" as const, content: SYSTEM_PROMPT },
      { role: "user" as const, content: buildUserPrompt({ intent, query: ctx.query, preview: ctx.preview }) },
    ],
  };

  const response = await client.chat(request);
  const content = typeof response?.message?.content === "string" ? response.message.content : "";
  const trimmed = content.trim();
  if (!trimmed) return [];
  if (trimmed === "AI unavailable.") {
    throw new Error(trimmed);
  }

  let parsed: unknown;
  try {
    parsed = parseJson(trimmed);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new Error(`AI returned invalid JSON: ${message}`);
  }

  const allowedColumns = new Set<string>((ctx.preview?.columns ?? []).map((c) => c.name));
  const ops = coerceOperations(parsed, allowedColumns);
  return ops;
}
