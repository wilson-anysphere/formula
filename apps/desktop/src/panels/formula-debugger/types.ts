export type TraceSpan = { start: number; end: number };

export type TraceReference =
  | { type: "cell"; cell: string }
  | { type: "range"; range: string };

export type TraceKind =
  | { type: "number" }
  | { type: "text" }
  | { type: "bool" }
  | { type: "cell_ref" }
  | { type: "range_ref" }
  | { type: "group" }
  | { type: "unary_op"; op: string }
  | { type: "binary_op"; op: string }
  | { type: "function_call"; name: string };

export type TraceValue =
  | null
  | number
  | string
  | boolean
  | { error: string }
  | { array: TraceValue[][] };

export type TraceNode = {
  kind: TraceKind;
  span: TraceSpan;
  value: TraceValue;
  reference?: TraceReference;
  children?: TraceNode[];
};

export type DebugStep = {
  id: string;
  text: string;
  span: TraceSpan;
  value: TraceValue;
  reference?: TraceReference;
  kind: TraceKind;
  children: DebugStep[];
};

