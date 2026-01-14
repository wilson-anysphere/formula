import { getActiveArgumentSpan } from "./activeArgument.js";

type FunctionCallContext = { name: string; argIndex: number };

export function getFunctionCallContext(formula: string, cursorIndex: number): FunctionCallContext | null {
  const active = getActiveArgumentSpan(formula, cursorIndex);
  if (!active) return null;
  return { name: active.fnName, argIndex: active.argIndex };
}
