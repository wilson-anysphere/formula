/**
 * OData query folding.
 *
 * This is a conservative folding engine that pushes a prefix of operations into
 * OData v4 query options (`$select`, `$filter`, `$orderby`, `$top`).
 *
 * Supported operations (prefix-only):
 *  - selectColumns -> $select
 *  - filterRows (simple comparisons) -> $filter
 *  - sortRows -> $orderby
 *  - take -> $top
 */

/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").FilterPredicate} FilterPredicate
 * @typedef {import("../model.js").ComparisonPredicate} ComparisonPredicate
 * @typedef {import("../model.js").SortSpec} SortSpec
 */

/**
 * @typedef {{
 *   select?: string[];
 *   filter?: string;
 *   orderby?: string;
 *   top?: number;
 * }} ODataQueryOptions
 */

/**
 * @typedef {{
 *   stepId: string;
 *   opType: QueryOperation["type"];
 *   status: "folded" | "local";
 *   url?: string;
 *   reason?: string;
 * }} ODataFoldingExplainStep
 */

/**
 * @typedef {{
 *   type: "local";
 *   url: string;
 *   query: ODataQueryOptions;
 * }} LocalPlan
 *
 * @typedef {{
 *   type: "odata";
 *   url: string;
 *   query: ODataQueryOptions;
 * }} ODataPlan
 *
 * @typedef {{
 *   type: "hybrid";
 *   url: string;
 *   query: ODataQueryOptions;
 *   localSteps: QueryStep[];
 * }} HybridPlan
 *
 * @typedef {LocalPlan | ODataPlan | HybridPlan} CompiledODataPlan
 */

/**
 * @typedef {{
 *   plan: CompiledODataPlan;
 *   steps: ODataFoldingExplainStep[];
 * }} ODataFoldingExplainResult
 */

/**
 * @param {string} input
 * @returns {boolean}
 */
function hasDuplicateStrings(input) {
  return new Set(input).size !== input.length;
}

/**
 * Extract a minimal set of OData query options from a URL.
 *
 * This allows `OData.Feed("...?$top=10")` style usage to participate in folding
 * without losing the user-supplied options.
 *
 * @param {string} url
 * @returns {ODataQueryOptions}
 */
function parseQueryOptionsFromUrl(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return {};
  }

  const params = parsed.searchParams;
  /** @type {ODataQueryOptions} */
  const out = {};

  const selectRaw = params.get("$select");
  if (typeof selectRaw === "string" && selectRaw.trim() !== "") {
    out.select = selectRaw
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0);
  }

  const filterRaw = params.get("$filter");
  if (typeof filterRaw === "string" && filterRaw.trim() !== "") {
    out.filter = filterRaw;
  }

  const orderbyRaw = params.get("$orderby");
  if (typeof orderbyRaw === "string" && orderbyRaw.trim() !== "") {
    out.orderby = orderbyRaw;
  }

  const topRaw = params.get("$top");
  if (typeof topRaw === "string" && topRaw.trim() !== "") {
    const parsedTop = Number.parseInt(topRaw, 10);
    if (Number.isFinite(parsedTop)) out.top = Math.max(0, parsedTop);
  }

  return out;
}

/**
 * Build a URL with OData query options appended.
 *
 * Note: We intentionally keep `$` unescaped in parameter names because that's
 * the conventional OData representation. Values are encoded with
 * `encodeURIComponent` (commas are left unescaped for readability).
 *
 * @param {string} baseUrl
 * @param {ODataQueryOptions | null | undefined} query
 * @returns {string}
 */
export function buildODataUrl(baseUrl, query) {
  const url = new URL(baseUrl);
  const normalized = query ?? {};

  const existingEntries = Array.from(url.searchParams.entries());
  /** @type {Map<string, string>} */
  const existing = new Map();
  for (const [k, v] of existingEntries) existing.set(k, v);

  // Only remove existing option keys when the caller provides overrides. This
  // keeps user-supplied query options (embedded in the base URL) intact for
  // non-folded requests.
  /** @type {Set<string>} */
  const overrideKeys = new Set();
  if (Array.isArray(normalized.select) && normalized.select.length > 0) overrideKeys.add("$select");
  if (typeof normalized.filter === "string" && normalized.filter.length > 0) overrideKeys.add("$filter");
  if (typeof normalized.orderby === "string" && normalized.orderby.length > 0) overrideKeys.add("$orderby");
  if (typeof normalized.top === "number" && Number.isFinite(normalized.top)) overrideKeys.add("$top");

  for (const key of overrideKeys) existing.delete(key);

  /** @type {Array<[string, string]>} */
  const entries = [];
  /** @type {Set<string>} */
  const seen = new Set();
  for (const [k] of existingEntries) {
    if (seen.has(k)) continue;
    seen.add(k);
    const v = existing.get(k);
    if (v != null) entries.push([k, v]);
  }

  if (Array.isArray(normalized.select) && normalized.select.length > 0) {
    entries.push(["$select", normalized.select.join(",")]);
  }
  if (typeof normalized.filter === "string" && normalized.filter.length > 0) {
    entries.push(["$filter", normalized.filter]);
  }
  if (typeof normalized.orderby === "string" && normalized.orderby.length > 0) {
    entries.push(["$orderby", normalized.orderby]);
  }
  if (typeof normalized.top === "number" && Number.isFinite(normalized.top)) {
    entries.push(["$top", String(Math.max(0, Math.trunc(normalized.top)))]);
  }

  const encodeKey = (k) => encodeURIComponent(k).replace(/%24/gi, "$");
  const encodeValue = (v) => encodeURIComponent(v).replace(/%2C/gi, ",");
  const qs = entries.map(([k, v]) => `${encodeKey(k)}=${encodeValue(v)}`).join("&");
  url.search = qs ? `?${qs}` : "";
  return url.toString();
}

/**
 * @param {unknown} value
 * @returns {string | null}
 */
function odataLiteral(value) {
  if (value == null) return "null";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : null;
  if (typeof value === "string") {
    // OData string literal escaping: single quotes are doubled.
    return `'${value.replaceAll("'", "''")}'`;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Best-effort: treat dates as ISO strings. Many OData services accept this
    // for Edm.DateTimeOffset comparisons.
    return `'${value.toISOString()}'`;
  }
  return null;
}

/**
 * Local predicate semantics stringify values (including null/undefined -> "") for
 * contains/startsWith/endsWith.
 *
 * @param {unknown} value
 * @returns {string}
 */
function valueToString(value) {
  if (value == null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString();
  return String(value);
}

/**
 * @param {ComparisonPredicate} predicate
 * @returns {string | null}
 */
function comparisonToFilter(predicate) {
  const col = predicate.column;

  const op = predicate.operator;
  if (op === "isNull") return `${col} eq null`;
  if (op === "isNotNull") return `${col} ne null`;

  switch (op) {
    case "equals":
    case "notEquals": {
      // Local semantics treat equals/notEquals as case-sensitive comparisons (even
      // when caseSensitive is provided), so do not apply `tolower()` here.
      const literal = odataLiteral(predicate.value);
      if (literal == null) return null;
      return op === "equals" ? `${col} eq ${literal}` : `${col} ne ${literal}`;
    }
    case "greaterThan":
    case "greaterThanOrEqual":
    case "lessThan":
    case "lessThanOrEqual": {
      // Local semantics return false when comparing against null/undefined.
      if (predicate.value == null) return "false";
      const literal = odataLiteral(predicate.value);
      if (literal == null) return null;
      return op === "greaterThan"
        ? `${col} gt ${literal}`
        : op === "greaterThanOrEqual"
          ? `${col} ge ${literal}`
          : op === "lessThan"
            ? `${col} lt ${literal}`
            : `${col} le ${literal}`;
    }
    case "contains":
    case "startsWith":
    case "endsWith": {
      const caseSensitive = predicate.caseSensitive ?? false;

      // Local semantics treat empty needle as always-true. OData `contains(null,'')`
      // can yield null/false, so avoid folding that case (run locally).
      const needleText = valueToString(predicate.value);
      if (needleText === "") return null;

      const needleLit = odataLiteral(needleText);
      if (needleLit == null) return null;

      // Best-effort: cast the column to Edm.String to mimic Power Query's local
      // `valueToString` behavior for these predicates.
      const textExpr = `cast(${col},Edm.String)`;
      const haystack = caseSensitive ? textExpr : `tolower(${textExpr})`;
      const needle = caseSensitive ? needleLit : `tolower(${needleLit})`;

      return op === "contains"
        ? `contains(${haystack}, ${needle})`
        : op === "startsWith"
          ? `startswith(${haystack}, ${needle})`
          : `endswith(${haystack}, ${needle})`;
    }
    default:
      return null;
  }
}

/**
 * @param {FilterPredicate} predicate
 * @returns {string | null}
 */
function predicateToFilter(predicate) {
  switch (predicate.type) {
    case "and": {
      if (predicate.predicates.length === 0) return "true";
      const parts = [];
      for (const p of predicate.predicates) {
        const compiled = predicateToFilter(p);
        if (!compiled) return null;
        parts.push(`(${compiled})`);
      }
      return parts.join(" and ");
    }
    case "or": {
      if (predicate.predicates.length === 0) return "false";
      const parts = [];
      for (const p of predicate.predicates) {
        const compiled = predicateToFilter(p);
        if (!compiled) return null;
        parts.push(`(${compiled})`);
      }
      return parts.join(" or ");
    }
    case "not": {
      const inner = predicateToFilter(predicate.predicate);
      if (!inner) return null;
      return `not (${inner})`;
    }
    case "comparison":
      return comparisonToFilter(predicate);
    default: {
      /** @type {never} */
      const exhausted = predicate;
      throw new Error(`Unsupported predicate '${exhausted.type}'`);
    }
  }
}

/**
 * @param {FilterPredicate} predicate
 * @returns {Set<string>}
 */
function collectPredicateColumns(predicate) {
  /** @type {Set<string>} */
  const cols = new Set();
  /**
   * @param {FilterPredicate} node
   */
  const visit = (node) => {
    switch (node.type) {
      case "comparison":
        cols.add(node.column);
        return;
      case "and":
      case "or":
        for (const child of node.predicates) visit(child);
        return;
      case "not":
        visit(node.predicate);
        return;
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported predicate '${exhausted.type}'`);
      }
    }
  };
  visit(predicate);
  return cols;
}

/**
 * @param {SortSpec[]} sortBy
 * @returns {string | null}
 */
function sortToOrderBy(sortBy) {
  if (!Array.isArray(sortBy) || sortBy.length === 0) return null;
  const parts = [];
  for (const spec of sortBy) {
    if (!spec || typeof spec !== "object") return null;
    if (spec.nulls != null) return null;
    const direction = spec.direction ?? "ascending";
    const dir = direction === "descending" ? "desc" : "asc";
    parts.push(`${spec.column} ${dir}`);
  }
  return parts.join(", ");
}

/**
 * @param {ODataQueryOptions} current
 * @param {QueryOperation} operation
 * @returns {ODataQueryOptions | null}
 */
function applyODataStep(current, operation) {
  switch (operation.type) {
    case "selectColumns": {
      if (!Array.isArray(operation.columns) || operation.columns.length === 0) return null;
      if (hasDuplicateStrings(operation.columns)) return null;
      if (Array.isArray(current.select) && current.select.length > 0) {
        const available = new Set(current.select);
        for (const col of operation.columns) {
          if (!available.has(col)) return null;
        }
      }
      return { ...current, select: operation.columns.slice() };
    }
    case "filterRows": {
      // OData query options apply `$filter` *before* `$top`. If a `$top` limit is
      // already in play (either from a previous `take` or embedded in the source
      // URL), folding a later `filterRows` step would change semantics:
      //   - local: take N rows, then filter
      //   - folded: filter all rows, then take N
      // Keep filtering local once `$top` has been introduced.
      if (typeof current.top === "number" && Number.isFinite(current.top)) return null;
      if (Array.isArray(current.select) && current.select.length > 0) {
        const available = new Set(current.select);
        for (const col of collectPredicateColumns(operation.predicate)) {
          if (!available.has(col)) return null;
        }
      }
      const compiled = predicateToFilter(operation.predicate);
      if (!compiled) return null;
      const prev = current.filter;
      const nextFilter = prev ? `(${prev}) and (${compiled})` : compiled;
      return { ...current, filter: nextFilter };
    }
    case "sortRows": {
      if (!Array.isArray(operation.sortBy)) return null;
      if (operation.sortBy.length === 0) return { ...current };
      // Similar to `$filter`, `$orderby` is applied before `$top` in OData.
      // Sorting after a `take` must stay local to preserve semantics.
      if (typeof current.top === "number" && Number.isFinite(current.top)) return null;
      if (Array.isArray(current.select) && current.select.length > 0) {
        const available = new Set(current.select);
        for (const spec of operation.sortBy) {
          if (!available.has(spec.column)) return null;
        }
      }
      const orderby = sortToOrderBy(operation.sortBy);
      if (!orderby) return null;
      return { ...current, orderby };
    }
    case "take": {
      const count = operation.count;
      if (typeof count !== "number" || !Number.isFinite(count)) return null;
      const next = Math.max(0, Math.trunc(count));
      const currentTop = typeof current.top === "number" && Number.isFinite(current.top) ? current.top : null;
      return { ...current, top: currentTop == null ? next : Math.min(currentTop, next) };
    }
    default:
      return null;
  }
}

/**
 * @param {QueryOperation} operation
 * @returns {string}
 */
function explainODataStepFailure(operation) {
  switch (operation.type) {
    case "selectColumns":
      return "invalid_select";
    case "filterRows":
      return "unsupported_predicate";
    case "sortRows":
      return "unsupported_sort";
    case "take":
      return "invalid_take";
    default:
      return "unsupported_op";
  }
}

export class ODataFoldingEngine {
  constructor() {
    /** @type {Set<QueryOperation["type"]>} */
    this.foldable = new Set(["selectColumns", "filterRows", "sortRows", "take"]);
  }

  /**
   * Explain folding decisions for an OData query.
   *
   * @param {Query} query
   * @returns {ODataFoldingExplainResult}
   */
  explain(query) {
    if (query.source.type !== "odata") {
      return {
        plan: { type: "local", url: "", query: {} },
        steps: (query.steps ?? []).map((step) => ({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: "unsupported_source",
        })),
      };
    }

    /** @type {ODataQueryOptions} */
    let current = parseQueryOptionsFromUrl(query.source.url);
    /** @type {QueryStep[]} */
    const localSteps = [];
    /** @type {ODataFoldingExplainStep[]} */
    const steps = [];
    let foldingBroken = false;
    let foldedCount = 0;

    for (const step of query.steps ?? []) {
      if (foldingBroken) {
        localSteps.push(step);
        steps.push({ stepId: step.id, opType: step.operation.type, status: "local", reason: "folding_stopped" });
        continue;
      }

      if (!this.foldable.has(step.operation.type)) {
        foldingBroken = true;
        localSteps.push(step);
        steps.push({ stepId: step.id, opType: step.operation.type, status: "local", reason: "unsupported_op" });
        continue;
      }

      const next = applyODataStep(current, step.operation);
      if (!next) {
        foldingBroken = true;
        localSteps.push(step);
        steps.push({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: explainODataStepFailure(step.operation),
        });
        continue;
      }

      current = next;
      foldedCount += 1;
      steps.push({
        stepId: step.id,
        opType: step.operation.type,
        status: "folded",
        url: buildODataUrl(query.source.url, current),
      });
    }

    const url = buildODataUrl(query.source.url, current);

    if (localSteps.length === 0) {
      return { plan: { type: "odata", url, query: current }, steps };
    }

    if (foldedCount === 0) {
      return { plan: { type: "local", url: query.source.url, query: {} }, steps };
    }

    return { plan: { type: "hybrid", url, query: current, localSteps }, steps };
  }
}
