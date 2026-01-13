import { getFunctionSignature, signatureParts, type FunctionSignature, type SignaturePart } from "./functionSignatures.js";
import { getActiveArgumentSpan } from "./activeArgument.js";

type FunctionCallContext = { name: string; argIndex: number };

export function getFunctionCallContext(formula: string, cursorIndex: number): FunctionCallContext | null {
  const active = getActiveArgumentSpan(formula, cursorIndex);
  if (!active) return null;
  return { name: active.fnName, argIndex: active.argIndex };
}

export type FunctionHint = {
  context: FunctionCallContext;
  signature: FunctionSignature;
  parts: SignaturePart[];
};

export function getFunctionHint(formula: string, cursorIndex: number): FunctionHint | null {
  const context = getFunctionCallContext(formula, cursorIndex);
  if (!context) return null;

  const signature = getFunctionSignature(context.name);
  if (!signature) return null;

  return {
    context,
    signature,
    parts: signatureParts(signature, context.argIndex),
  };
}
