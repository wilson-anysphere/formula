/**
 * Tool result serialization for feeding results back into the LLM context.
 *
 * Why this exists:
 * - Some tools (notably `read_range`) can return huge matrices.
 * - Appending the full JSON into the conversation can blow prompt context limits
 *   and balloon any persisted audit logs.
 *
 * This module provides bounded, per-tool summaries intended for LLM consumption.
 * The full tool result object is still available to callers (e.g. audit hooks).
 */
 
/**
 * @typedef {import("./types.js").ToolCall} ToolCall
 */
 
const DEFAULT_MAX_CHARS = 20_000;
const CHARS_PER_TOKEN_APPROX = 4;

/**
 * Tool names are identifiers and should not include leading/trailing whitespace.
 *
 * @param {unknown} value
 * @param {string} fallback
 * @returns {string}
 */
function normalizeToolName(value, fallback) {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim();
  return trimmed ? trimmed : fallback;
}

/**
 * @param {{ toolCall: ToolCall, result: unknown, maxChars?: number, maxTokens?: number }} params
 * @returns {string}
 */
export function serializeToolResultForModel(params) {
  const maxChars = resolveMaxChars(params);
  const toolName = normalizeToolName(params?.toolCall?.name, "");

  // Prefer deterministic, per-tool summaries for high-volume tools.
  if (toolName === "read_range") {
    return serializeReadRange({ toolCall: params.toolCall, result: params.result, maxChars });
  }
  if (toolName === "filter_range") {
    return serializeFilterRange({ toolCall: params.toolCall, result: params.result, maxChars });
  }
  if (toolName === "detect_anomalies") {
    return serializeDetectAnomalies({ toolCall: params.toolCall, result: params.result, maxChars });
  }

  return serializeGeneric({ toolCall: params.toolCall, result: params.result, maxChars });
}
 
/**
 * @param {{ maxChars?: number, maxTokens?: number }} params
 */
function resolveMaxChars(params) {
  if (typeof params.maxChars === "number" && Number.isFinite(params.maxChars) && params.maxChars > 0) {
    return Math.floor(params.maxChars);
  }
  if (typeof params.maxTokens === "number" && Number.isFinite(params.maxTokens) && params.maxTokens > 0) {
    return Math.floor(params.maxTokens * CHARS_PER_TOKEN_APPROX);
  }
  return DEFAULT_MAX_CHARS;
}
 
/**
 * @param {{ toolCall: ToolCall, result: unknown, maxChars: number }} params
 * @returns {string}
 */
function serializeReadRange(params) {
  const base = normalizeToolExecutionEnvelope(params.toolCall, params.result);
  const tool = String(typeof base.tool === "string" && base.tool.trim() ? base.tool : String(params.toolCall?.name ?? "read_range"));
  const data = base.data && typeof base.data === "object" ? base.data : null;
  const values = /** @type {unknown} */ (data?.values);
  const formulas = /** @type {unknown} */ (data?.formulas);
  const range = typeof data?.range === "string" ? data.range : safeRangeFromCall(params.toolCall);
  const toolTruncated = typeof data?.truncated === "boolean" ? data.truncated : false;
  
  if (!Array.isArray(values)) {
    return serializeGeneric(params);
  }
  
  const rows = values.length;
  const cols = Array.isArray(values[0]) ? values[0].length : 0;
 
  // We iteratively shrink the preview until we fit the maxChars budget.
  const attempts = [
    { rows: 20, cols: 10, cellStringChars: 200 },
    { rows: 10, cols: 8, cellStringChars: 120 },
    { rows: 5, cols: 5, cellStringChars: 80 },
    { rows: 3, cols: 3, cellStringChars: 60 }
  ];
 
  for (const attempt of attempts) {
    const previewRows = Math.min(rows, attempt.rows);
    const previewCols = Math.min(cols, attempt.cols);
    const previewValues = sliceMatrix(values, previewRows, previewCols, attempt.cellStringChars);
    const previewFormulas = Array.isArray(formulas)
      ? sliceMatrix(formulas, previewRows, previewCols, attempt.cellStringChars)
      : undefined;
  
    const truncated = toolTruncated || rows > previewRows || cols > previewCols;
  
    const payload = {
      ...base,
      tool,
      data: {
        range,
        shape: { rows, cols },
        values: previewValues,
        ...(previewFormulas ? { formulas: previewFormulas } : {}),
        truncated,
        truncated_rows: Math.max(0, rows - previewRows),
        truncated_cols: Math.max(0, cols - previewCols)
      }
    };
 
    const json = safeJsonStringify(payload);
    if (json.length <= params.maxChars) return json;
  }
  
  // Fallback: minimal summary that always fits.
  return finalizeJson(
    safeJsonStringify({ ...base, tool, data: { range, shape: { rows, cols }, truncated: true } }),
    params.maxChars,
    {
      tool,
      ok: base.ok,
      ...(base.timing ? { timing: base.timing } : {}),
      ...(base.error ? { error: base.error } : {}),
      data: { range, shape: { rows, cols }, truncated: true }
    }
  );
}
 
/**
 * @param {{ toolCall: ToolCall, result: unknown, maxChars: number }} params
 * @returns {string}
 */
function serializeFilterRange(params) {
  const base = normalizeToolExecutionEnvelope(params.toolCall, params.result);
  const tool = String(typeof base.tool === "string" && base.tool.trim() ? base.tool : String(params.toolCall?.name ?? "filter_range"));
  const data = base.data && typeof base.data === "object" ? base.data : null;
  const range = typeof data?.range === "string" ? data.range : safeRangeFromCall(params.toolCall);
  const rows = Array.isArray(data?.matching_rows) ? data.matching_rows : null;
  const count = typeof data?.count === "number" ? data.count : rows ? rows.length : undefined;
  const toolTruncated = typeof data?.truncated === "boolean" ? data.truncated : false;
  
  const attempts = [200, 100, 50, 20];
  for (const limit of attempts) {
    const list = rows ? rows.slice(0, limit) : undefined;
    const previewCount = list ? list.length : 0;
    const truncated =
      toolTruncated ||
      (rows ? rows.length > previewCount : false) ||
      (typeof count === "number" ? count > previewCount : false);
  
    const payload = {
      ...base,
      tool,
      data: {
        ...(range ? { range } : {}),
        ...(typeof count === "number" ? { count } : {}),
        ...(list ? { matching_rows: list } : {}),
        ...(rows || toolTruncated ? { truncated } : {})
      }
    };
    const json = safeJsonStringify(payload);
    if (json.length <= params.maxChars) return json;
  }
  
  return finalizeJson(
    safeJsonStringify({
      ...base,
      tool,
      data: { ...(range ? { range } : {}), ...(typeof count === "number" ? { count } : {}), truncated: true }
    }),
    params.maxChars,
    {
      tool,
      ok: base.ok,
      ...(base.timing ? { timing: base.timing } : {}),
      ...(base.error ? { error: base.error } : {}),
      data: { ...(range ? { range } : {}), ...(typeof count === "number" ? { count } : {}), truncated: true }
    }
  );
}

/**
 * @param {{ toolCall: ToolCall, result: unknown, maxChars: number }} params
 * @returns {string}
 */
function serializeDetectAnomalies(params) {
  const base = normalizeToolExecutionEnvelope(params.toolCall, params.result);
  const tool = String(
    typeof base.tool === "string" && base.tool.trim() ? base.tool : String(params.toolCall?.name ?? "detect_anomalies")
  );
  const data = base.data && typeof base.data === "object" ? base.data : null;
  const range = typeof data?.range === "string" ? data.range : safeRangeFromCall(params.toolCall);
  const method =
    typeof data?.method === "string"
      ? data.method
      : (() => {
          const args = params.toolCall?.arguments;
          if (args && typeof args === "object" && !Array.isArray(args) && typeof args.method === "string") return args.method;
          return undefined;
        })();

  const anomalies = Array.isArray(data?.anomalies) ? data.anomalies : null;
  const toolTruncated = typeof data?.truncated === "boolean" ? data.truncated : false;

  // Prefer explicit total_anomalies when the tool has already truncated the list.
  const totalAnomalies =
    typeof data?.total_anomalies === "number"
      ? data.total_anomalies
      : typeof data?.count === "number"
        ? data.count
        : anomalies
          ? anomalies.length
          : undefined;

  const attempts = [200, 100, 50, 20, 10, 5];
  for (const limit of attempts) {
    const preview = anomalies
      ? anomalies.slice(0, limit).map((entry) => {
          if (typeof entry !== "object" || entry === null) return truncatePrimitive(entry, 200);
          const record = /** @type {any} */ (entry);
          const out = {
            ...(typeof record.cell === "string" ? { cell: truncateString(record.cell, 200) } : {}),
            ...("value" in record ? { value: truncatePrimitive(record.value ?? null, 200) } : {}),
            ...("score" in record ? { score: truncatePrimitive(record.score ?? null, 200) } : {})
          };
          // If we couldn't extract known fields, fall back to a truncated string representation.
          if (!("cell" in out) && !("value" in out) && !("score" in out)) {
            return truncatePrimitive(record, 200);
          }
          return out;
        })
      : undefined;

    const previewCount = preview ? preview.length : 0;
    const truncated =
      toolTruncated ||
      (anomalies ? anomalies.length > previewCount : false) ||
      (typeof totalAnomalies === "number" ? totalAnomalies > previewCount : false);

    const payload = {
      ...base,
      tool,
      data: {
        ...(range ? { range } : {}),
        ...(method ? { method } : {}),
        ...(typeof totalAnomalies === "number"
          ? typeof data?.total_anomalies === "number"
            ? { total_anomalies: totalAnomalies }
            : { count: totalAnomalies }
          : {}),
        ...(preview ? { anomalies: preview } : {}),
        ...(anomalies || toolTruncated ? { truncated } : {})
      }
    };

    const json = safeJsonStringify(payload);
    if (json.length <= params.maxChars) return json;
  }

  return finalizeJson(
    safeJsonStringify({
      ...base,
      tool,
      data: {
        ...(range ? { range } : {}),
        ...(method ? { method } : {}),
        ...(typeof totalAnomalies === "number"
          ? typeof data?.total_anomalies === "number"
            ? { total_anomalies: totalAnomalies }
            : { count: totalAnomalies }
          : {}),
        truncated: true
      }
    }),
    params.maxChars,
    {
      tool,
      ok: base.ok,
      ...(base.timing ? { timing: base.timing } : {}),
      ...(base.error ? { error: base.error } : {}),
      data: {
        ...(range ? { range } : {}),
        ...(method ? { method } : {}),
        ...(typeof totalAnomalies === "number"
          ? typeof data?.total_anomalies === "number"
            ? { total_anomalies: totalAnomalies }
            : { count: totalAnomalies }
          : {}),
        truncated: true
      }
    }
  );
}

/**
 * Generic fallback: deep truncation of arbitrary results.
 *
 * @param {{ toolCall: ToolCall, result: unknown, maxChars: number }} params
 * @returns {string}
 */
function serializeGeneric(params) {
  const base = normalizeToolExecutionEnvelope(params.toolCall, params.result);
 
  const attempts = [
    { depth: 6, maxArray: 100, maxKeys: 100, maxString: 2_000 },
    { depth: 5, maxArray: 50, maxKeys: 50, maxString: 1_000 },
    { depth: 4, maxArray: 25, maxKeys: 25, maxString: 500 },
    { depth: 3, maxArray: 10, maxKeys: 10, maxString: 200 }
  ];
 
  for (const attempt of attempts) {
    const truncated = truncateUnknown(base, {
      maxDepth: attempt.depth,
      maxArrayLength: attempt.maxArray,
      maxObjectKeys: attempt.maxKeys,
      maxStringLength: attempt.maxString
    });
    const json = safeJsonStringify(truncated.value);
    if (json.length <= params.maxChars) return json;
  }
 
  return finalizeJson(
    safeJsonStringify({
      tool: normalizeToolName(params.toolCall?.name, "tool"),
      ok: typeof base.ok === "boolean" ? base.ok : undefined,
      truncated: true,
      note: "Tool result exceeded max serialization budget."
    }),
    params.maxChars,
    { tool: normalizeToolName(params.toolCall?.name, "tool"), truncated: true }
  );
}
 
/**
 * @param {ToolCall} toolCall
 * @param {unknown} result
 */
function normalizeToolExecutionEnvelope(toolCall, result) {
  const toolName = normalizeToolName(toolCall?.name, "tool");

  if (result && typeof result === "object" && !Array.isArray(result)) {
    // If this is already a ToolExecutionResult-like envelope, keep it.
    if ("tool" in result || "ok" in result || "data" in result || "error" in result) {
      const envelope = /** @type {any} */ (result);
      const tool = normalizeToolName(envelope.tool, toolName);
      if (tool !== envelope.tool) return { ...envelope, tool };
      return envelope;
    }
  }

  // Otherwise wrap in a generic envelope.
  return {
    tool: toolName,
    ok: true,
    data: result
  };
}
 
/**
 * @param {ToolCall} toolCall
 */
function safeRangeFromCall(toolCall) {
  const args = toolCall?.arguments;
  if (args && typeof args === "object" && !Array.isArray(args) && typeof args.range === "string") {
    return args.range;
  }
  return undefined;
}
 
/**
 * @param {unknown} matrix
 * @param {number} maxRows
 * @param {number} maxCols
 * @param {number} maxStringLength
 */
function sliceMatrix(matrix, maxRows, maxCols, maxStringLength) {
  if (!Array.isArray(matrix)) return [];
  const out = [];
  const rows = Math.min(matrix.length, maxRows);
  for (let r = 0; r < rows; r++) {
    const row = matrix[r];
    if (!Array.isArray(row)) {
      out.push([]);
      continue;
    }
    const cols = Math.min(row.length, maxCols);
    const nextRow = [];
    for (let c = 0; c < cols; c++) {
      nextRow.push(truncatePrimitive(row[c], maxStringLength));
    }
    out.push(nextRow);
  }
  return out;
}
 
/**
 * @param {unknown} value
 * @param {number} maxStringLength
 */
function truncatePrimitive(value, maxStringLength) {
  if (typeof value === "string") return truncateString(value, maxStringLength);
  if (typeof value === "number" || typeof value === "boolean" || value === null) return value;
  // Defensive: tool matrices should be scalar, but keep it safe.
  if (value === undefined) return null;
  return truncateString(safeJsonStringify(value), maxStringLength);
}
 
/**
 * @param {string} value
 * @param {number} maxLength
 */
function truncateString(value, maxLength) {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, maxLength)}…[truncated ${value.length - maxLength} chars]`;
}
 
/**
 * @param {unknown} value
 * @returns {string}
 */
function safeJsonStringify(value) {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    try {
      return JSON.stringify(String(value));
    } catch {
      return String(value);
    }
  }
}

/**
 * Ensure the returned string is within budget. Prefer returning valid JSON when possible.
 *
 * @param {string} json
 * @param {number} maxChars
 * @param {unknown} fallbackValue
 */
function finalizeJson(json, maxChars, fallbackValue) {
  if (json.length <= maxChars) return json;
  const fallback = safeJsonStringify(fallbackValue);
  if (fallback.length <= maxChars) return fallback;
  if (maxChars <= 0) return "";
  if (maxChars < 2) return "";
  // Absolute last resort: return a tiny JSON object.
  return "{}";
}
 
/**
 * @typedef {{
 *   maxDepth: number,
 *   maxArrayLength: number,
 *   maxObjectKeys: number,
 *   maxStringLength: number
 * }} TruncateOptions
 */
 
/**
 * @param {unknown} value
 * @param {TruncateOptions} options
 * @returns {{ value: unknown, truncated: boolean }}
 */
function truncateUnknown(value, options) {
  const seen = new WeakSet();
  const walk = (v, depth) => {
    if (typeof v === "string") {
      const truncated = v.length > options.maxStringLength;
      return { value: truncateString(v, options.maxStringLength), truncated };
    }
    if (typeof v === "number" || typeof v === "boolean" || v === null) return { value: v, truncated: false };
    if (v === undefined) return { value: null, truncated: true };
 
    if (depth >= options.maxDepth) {
      return { value: "[truncated: max depth]", truncated: true };
    }
 
    if (Array.isArray(v)) {
      const out = [];
      let truncated = false;
      const len = Math.min(v.length, options.maxArrayLength);
      for (let i = 0; i < len; i++) {
        const child = walk(v[i], depth + 1);
        out.push(child.value);
        truncated ||= child.truncated;
      }
      if (v.length > len) {
        out.push(`[truncated: ${v.length - len} more items]`);
        truncated = true;
      }
      return { value: out, truncated };
    }
 
    if (typeof v === "object") {
      if (seen.has(/** @type {object} */ (v))) return { value: "[truncated: circular]", truncated: true };
      seen.add(/** @type {object} */ (v));
 
      const obj = /** @type {Record<string, unknown>} */ (v);
      const keys = Object.keys(obj);
      const out = {};
      let truncated = false;
      const len = Math.min(keys.length, options.maxObjectKeys);
      for (let i = 0; i < len; i++) {
        const key = keys[i];
        const child = walk(obj[key], depth + 1);
        out[key] = child.value;
        truncated ||= child.truncated;
      }
      if (keys.length > len) {
        out.__truncated_keys__ = keys.length - len;
        truncated = true;
      }
      return { value: out, truncated };
    }
 
    // bigint/symbol/function/etc
    return { value: truncateString(String(v), options.maxStringLength), truncated: true };
  };
 
  return walk(value, 0);
}
