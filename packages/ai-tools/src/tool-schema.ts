import { z } from "zod";
import { columnLabelToIndex, parseA1Cell, parseA1Range, rangeSize } from "./spreadsheet/a1.js";
import type { CellScalar } from "./spreadsheet/types.js";

export type ToolName =
  | "read_range"
  | "write_cell"
  | "set_range"
  | "apply_formula_column"
  | "create_pivot_table"
  | "create_chart"
  | "sort_range"
  | "filter_range"
  | "apply_formatting"
  | "detect_anomalies"
  | "compute_statistics"
  | "fetch_external_data";

export const ToolNameSchema = z.enum([
  "read_range",
  "write_cell",
  "set_range",
  "apply_formula_column",
  "create_pivot_table",
  "create_chart",
  "sort_range",
  "filter_range",
  "apply_formatting",
  "detect_anomalies",
  "compute_statistics",
  "fetch_external_data"
]);

const CellScalarSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const A1CellSchema = z.string().min(1).superRefine((value, ctx) => {
  try {
    parseA1Cell(value);
  } catch (error) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: error instanceof Error ? error.message : `Invalid A1 cell reference: ${value}`
    });
  }
});

const A1RangeSchema = z.string().min(1).superRefine((value, ctx) => {
  try {
    parseA1Range(value);
  } catch (error) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: error instanceof Error ? error.message : `Invalid A1 range reference: ${value}`
    });
  }
});

const ColumnSchema = z.string().min(1).superRefine((value, ctx) => {
  try {
    columnLabelToIndex(value);
  } catch (error) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: error instanceof Error ? error.message : `Invalid column label: ${value}`
    });
  }
});

export const ReadRangeParamsSchema = z.object({
  range: A1RangeSchema,
  include_formulas: z.boolean().optional().default(false)
});

export type ReadRangeParams = z.infer<typeof ReadRangeParamsSchema>;

export const WriteCellParamsSchema = z.object({
  cell: A1CellSchema,
  value: CellScalarSchema,
  is_formula: z.boolean().optional()
});

export type WriteCellParams = z.infer<typeof WriteCellParamsSchema>;

export const SetRangeParamsSchema = z
  .object({
    range: A1RangeSchema,
    values: z.array(z.array(CellScalarSchema)),
    interpret_as: z.enum(["auto", "value", "formula"]).optional().default("auto")
  })
  .superRefine((data, ctx) => {
    const size = rangeSize(parseA1Range(data.range));
    if (data.values.length !== size.rows) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `set_range values expected ${size.rows} rows but got ${data.values.length}`
      });
      return;
    }
    for (const [rowIndex, row] of data.values.entries()) {
      if (row.length !== size.cols) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `set_range values row ${rowIndex} expected ${size.cols} columns but got ${row.length}`
        });
      }
    }
  });

export type SetRangeParams = z.infer<typeof SetRangeParamsSchema>;

export const ApplyFormulaColumnParamsSchema = z.object({
  column: ColumnSchema,
  formula_template: z.string().min(1),
  start_row: z.number().int().positive(),
  end_row: z.number().int().optional().default(-1)
});

export type ApplyFormulaColumnParams = z.infer<typeof ApplyFormulaColumnParamsSchema>;

const AggregationSchema = z.preprocess(
  (value) => (typeof value === "string" ? value.toLowerCase() : value),
  z.enum(["sum", "count", "average", "max", "min"])
);

export const CreatePivotTableParamsSchema = z.object({
  source_range: A1RangeSchema,
  rows: z.array(z.string().min(1)).min(1),
  columns: z.array(z.string().min(1)).optional(),
  values: z
    .array(
      z.object({
        field: z.string().min(1),
        aggregation: AggregationSchema
      })
    )
    .min(1),
  destination: z.string().min(1)
});

export type CreatePivotTableParams = z.infer<typeof CreatePivotTableParamsSchema>;

export const CreateChartParamsSchema = z.object({
  chart_type: z.enum(["bar", "line", "pie", "scatter", "area"]),
  data_range: A1RangeSchema,
  title: z.string().optional(),
  position: z.string().optional()
});

export type CreateChartParams = z.infer<typeof CreateChartParamsSchema>;

export const SortRangeParamsSchema = z.object({
  range: A1RangeSchema,
  sort_by: z
    .array(
      z.object({
        column: ColumnSchema,
        order: z.enum(["asc", "desc"]).default("asc")
      })
    )
    .min(1),
  has_header: z.boolean().optional().default(false)
});

export type SortRangeParams = z.infer<typeof SortRangeParamsSchema>;

export const FilterRangeParamsSchema = z.object({
  range: A1RangeSchema,
  criteria: z
    .array(
      z.object({
        column: ColumnSchema,
        operator: z.enum(["equals", "contains", "greater", "less", "between"]),
        value: z.union([z.string(), z.number()]),
        value2: z.union([z.string(), z.number()]).optional()
      })
    )
    .min(1),
  has_header: z.boolean().optional().default(false)
});

export type FilterRangeParams = z.infer<typeof FilterRangeParamsSchema>;

export const ApplyFormattingParamsSchema = z.object({
  range: A1RangeSchema,
  format: z
    .object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      font_size: z.number().int().positive().optional(),
      font_color: z.string().optional(),
      background_color: z.string().optional(),
      number_format: z.string().optional(),
      horizontal_align: z.enum(["left", "center", "right"]).optional()
    })
    .refine((format) => Object.keys(format).length > 0, "format must specify at least one field")
});

export type ApplyFormattingParams = z.infer<typeof ApplyFormattingParamsSchema>;

export const DetectAnomaliesParamsSchema = z.object({
  range: A1RangeSchema,
  method: z.enum(["zscore", "iqr", "isolation_forest"]).optional().default("zscore"),
  threshold: z.number().positive().optional()
});

export type DetectAnomaliesParams = z.infer<typeof DetectAnomaliesParamsSchema>;

export const ComputeStatisticsParamsSchema = z.object({
  range: A1RangeSchema,
  measures: z
    .array(z.enum(["mean", "median", "mode", "stdev", "variance", "min", "max", "quartiles", "correlation"]))
    .optional()
    .default(["mean", "median", "stdev", "min", "max"])
});

export type ComputeStatisticsParams = z.infer<typeof ComputeStatisticsParamsSchema>;

export const FetchExternalDataParamsSchema = z.object({
  source_type: z.enum(["api"]),
  url: z.string().url(),
  destination: A1CellSchema,
  transform: z.enum(["json_to_table", "raw_text"]).optional().default("json_to_table"),
  headers: z.record(z.string()).optional()
});

export type FetchExternalDataParams = z.infer<typeof FetchExternalDataParamsSchema>;

export type ToolParamsByName = {
  read_range: ReadRangeParams;
  write_cell: WriteCellParams;
  set_range: SetRangeParams;
  apply_formula_column: ApplyFormulaColumnParams;
  create_pivot_table: CreatePivotTableParams;
  create_chart: CreateChartParams;
  sort_range: SortRangeParams;
  filter_range: FilterRangeParams;
  apply_formatting: ApplyFormattingParams;
  detect_anomalies: DetectAnomaliesParams;
  compute_statistics: ComputeStatisticsParams;
  fetch_external_data: FetchExternalDataParams;
};

export interface ToolCall<TName extends ToolName = ToolName> {
  name: TName;
  parameters: ToolParamsByName[TName];
}

export interface UnknownToolCall {
  name: string;
  parameters: unknown;
}

export interface ToolDefinition {
  name: ToolName;
  description: string;
  parameters: Record<string, unknown>;
}

export interface ToolRegistryEntry<TName extends ToolName> {
  name: TName;
  description: string;
  paramsSchema: z.ZodType<ToolParamsByName[TName], z.ZodTypeDef, unknown>;
  jsonSchema: Record<string, unknown>;
}

export const TOOL_REGISTRY: { [K in ToolName]: ToolRegistryEntry<K> } = {
  read_range: {
    name: "read_range",
    description: "Read cell values from a range",
    paramsSchema: ReadRangeParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range in A1 notation (e.g., 'Sheet1!A1:D10')" },
        include_formulas: { type: "boolean", default: false }
      },
      required: ["range"]
    }
  },
  write_cell: {
    name: "write_cell",
    description: "Write a value or formula to a cell",
    paramsSchema: WriteCellParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        cell: { type: "string", description: "Cell reference (e.g., 'Sheet1!A1')" },
        value: { description: "Scalar value or formula string", anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }, { type: "null" }] },
        is_formula: { type: "boolean", description: "Treat value as formula even if it does not start with '='." }
      },
      required: ["cell", "value"]
    }
  },
  set_range: {
    name: "set_range",
    description: "Set a rectangular range of values/formulas in one operation",
    paramsSchema: SetRangeParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range in A1 notation (e.g., 'Sheet1!A1:B3')" },
        values: {
          type: "array",
          items: {
            type: "array",
            items: { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }, { type: "null" }] }
          }
        },
        interpret_as: { type: "string", enum: ["auto", "value", "formula"], default: "auto" }
      },
      required: ["range", "values"]
    }
  },
  apply_formula_column: {
    name: "apply_formula_column",
    description: "Apply a formula template with a {row} placeholder to a column.",
    paramsSchema: ApplyFormulaColumnParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        column: { type: "string", description: "Column label (e.g., 'C')" },
        formula_template: { type: "string", description: "Formula with {row} placeholder (e.g., '=A{row}*B{row}')" },
        start_row: { type: "number" },
        end_row: { type: "number", description: "-1 means last used row on the sheet" }
      },
      required: ["column", "formula_template", "start_row"]
    }
  },
  create_pivot_table: {
    name: "create_pivot_table",
    description: "Create a pivot table from a source range (stub).",
    paramsSchema: CreatePivotTableParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        source_range: { type: "string" },
        rows: { type: "array", items: { type: "string" } },
        columns: { type: "array", items: { type: "string" } },
        values: {
          type: "array",
          items: {
            type: "object",
            properties: {
              field: { type: "string" },
              aggregation: { type: "string", enum: ["sum", "count", "average", "max", "min"] }
            },
            required: ["field", "aggregation"]
          }
        },
        destination: { type: "string" }
      },
      required: ["source_range", "rows", "values"]
    }
  },
  create_chart: {
    name: "create_chart",
    description: "Create a chart from a data range (stub).",
    paramsSchema: CreateChartParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        chart_type: { type: "string", enum: ["bar", "line", "pie", "scatter", "area"] },
        data_range: { type: "string" },
        title: { type: "string" },
        position: { type: "string" }
      },
      required: ["chart_type", "data_range"]
    }
  },
  sort_range: {
    name: "sort_range",
    description: "Sort a range by one or more columns.",
    paramsSchema: SortRangeParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string" },
        sort_by: {
          type: "array",
          items: {
            type: "object",
            properties: {
              column: { type: "string" },
              order: { type: "string", enum: ["asc", "desc"] }
            },
            required: ["column"]
          }
        },
        has_header: { type: "boolean", default: false }
      },
      required: ["range", "sort_by"]
    }
  },
  filter_range: {
    name: "filter_range",
    description: "Filter a range based on column criteria (does not modify cells; returns matching rows).",
    paramsSchema: FilterRangeParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string" },
        criteria: {
          type: "array",
          items: {
            type: "object",
            properties: {
              column: { type: "string" },
              operator: { type: "string", enum: ["equals", "contains", "greater", "less", "between"] },
              value: { anyOf: [{ type: "string" }, { type: "number" }] },
              value2: { anyOf: [{ type: "string" }, { type: "number" }] }
            },
            required: ["column", "operator", "value"]
          }
        },
        has_header: { type: "boolean", default: false }
      },
      required: ["range", "criteria"]
    }
  },
  apply_formatting: {
    name: "apply_formatting",
    description: "Apply formatting attributes to a range.",
    paramsSchema: ApplyFormattingParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string" },
        format: {
          type: "object",
          properties: {
            bold: { type: "boolean" },
            italic: { type: "boolean" },
            font_size: { type: "number" },
            font_color: { type: "string" },
            background_color: { type: "string" },
            number_format: { type: "string" },
            horizontal_align: { type: "string", enum: ["left", "center", "right"] }
          }
        }
      },
      required: ["range", "format"]
    }
  },
  detect_anomalies: {
    name: "detect_anomalies",
    description: "Find outliers and anomalies in a range.",
    paramsSchema: DetectAnomaliesParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string" },
        method: { type: "string", enum: ["zscore", "iqr", "isolation_forest"], default: "zscore" },
        threshold: { type: "number" }
      },
      required: ["range"]
    }
  },
  compute_statistics: {
    name: "compute_statistics",
    description: "Compute descriptive statistics for a range.",
    paramsSchema: ComputeStatisticsParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        range: { type: "string" },
        measures: {
          type: "array",
          items: {
            type: "string",
            enum: ["mean", "median", "mode", "stdev", "variance", "min", "max", "quartiles", "correlation"]
          }
        }
      },
      required: ["range"]
    }
  },
  fetch_external_data: {
    name: "fetch_external_data",
    description: "Fetch external data from an API and write it into the sheet (guarded).",
    paramsSchema: FetchExternalDataParamsSchema,
    jsonSchema: {
      type: "object",
      properties: {
        source_type: { type: "string", enum: ["api"] },
        url: { type: "string" },
        destination: { type: "string", description: "Top-left cell to write the fetched data to." },
        transform: { type: "string", enum: ["json_to_table", "raw_text"], default: "json_to_table" },
        headers: { type: "object", additionalProperties: { type: "string" } }
      },
      required: ["source_type", "url", "destination"]
    }
  }
};

export const SPREADSHEET_TOOL_DEFINITIONS: ToolDefinition[] = Object.values(TOOL_REGISTRY).map((tool) => ({
  name: tool.name,
  description: tool.description,
  parameters: tool.jsonSchema
}));

export function validateToolCall(call: UnknownToolCall): ToolCall {
  const name = ToolNameSchema.parse(call.name);
  const entry = TOOL_REGISTRY[name];
  const normalized = normalizeToolParameters(name, call.parameters);
  const parameters = entry.paramsSchema.parse(normalized);
  return { name, parameters } as ToolCall;
}

function normalizeToolParameters(name: ToolName, parameters: unknown): unknown {
  if (!parameters || typeof parameters !== "object" || Array.isArray(parameters)) return parameters;

  const params = { ...(parameters as Record<string, unknown>) } as Record<string, any>;

  switch (name) {
    case "read_range":
      if (params.include_formulas === undefined && params.includeFormulas !== undefined) {
        params.include_formulas = params.includeFormulas;
      }
      break;
    case "write_cell":
      if (params.is_formula === undefined && params.isFormula !== undefined) {
        params.is_formula = params.isFormula;
      }
      break;
    case "set_range":
      if (params.interpret_as === undefined && params.interpretAs !== undefined) {
        params.interpret_as = params.interpretAs;
      }
      break;
    case "apply_formula_column":
      if (params.formula_template === undefined && params.formulaTemplate !== undefined) {
        params.formula_template = params.formulaTemplate;
      }
      if (params.start_row === undefined && params.startRow !== undefined) {
        params.start_row = params.startRow;
      }
      if (params.end_row === undefined && params.endRow !== undefined) {
        params.end_row = params.endRow;
      }
      break;
    case "create_pivot_table":
      if (params.source_range === undefined && params.sourceRange !== undefined) {
        params.source_range = params.sourceRange;
      }
      break;
    case "create_chart":
      if (params.chart_type === undefined && params.chartType !== undefined) {
        params.chart_type = params.chartType;
      }
      if (params.data_range === undefined && params.dataRange !== undefined) {
        params.data_range = params.dataRange;
      }
      break;
    case "sort_range":
      if (params.sort_by === undefined && params.sortBy !== undefined) {
        params.sort_by = params.sortBy;
      }
      if (params.has_header === undefined && params.hasHeader !== undefined) {
        params.has_header = params.hasHeader;
      }
      break;
    case "filter_range":
      if (params.has_header === undefined && params.hasHeader !== undefined) {
        params.has_header = params.hasHeader;
      }
      break;
    case "apply_formatting":
      if (params.format && typeof params.format === "object" && !Array.isArray(params.format)) {
        const fmt = { ...(params.format as Record<string, unknown>) } as Record<string, any>;
        if (fmt.font_size === undefined && fmt.fontSize !== undefined) fmt.font_size = fmt.fontSize;
        if (fmt.font_color === undefined && fmt.fontColor !== undefined) fmt.font_color = fmt.fontColor;
        if (fmt.background_color === undefined && fmt.backgroundColor !== undefined) fmt.background_color = fmt.backgroundColor;
        if (fmt.number_format === undefined && fmt.numberFormat !== undefined) fmt.number_format = fmt.numberFormat;
        if (fmt.horizontal_align === undefined && fmt.horizontalAlign !== undefined) fmt.horizontal_align = fmt.horizontalAlign;
        params.format = fmt;
      }
      break;
    case "detect_anomalies":
      // No aliases currently.
      break;
    case "compute_statistics":
      // No aliases currently.
      break;
    case "fetch_external_data":
      if (params.source_type === undefined && params.sourceType !== undefined) {
        params.source_type = params.sourceType;
      }
      break;
    default: {
      const exhaustive: never = name;
      return parameters;
    }
  }

  return params;
}
