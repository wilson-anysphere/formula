import { getFunctionSignature, signatureParts, type FunctionSignature, type SignaturePart } from "./functionSignatures.js";

export type FunctionCallContext = { name: string; argIndex: number };

type StackFrame = { kind: "function"; name: string; argIndex: number } | { kind: "group" };

function isIdentifierStart(ch: string): boolean {
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z") || ch === "_";
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || (ch >= "0" && ch <= "9") || ch === ".";
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

export function getFunctionCallContext(formula: string, cursorIndex: number): FunctionCallContext | null {
  const cursor = Math.max(0, Math.min(cursorIndex, formula.length));
  const stack: StackFrame[] = [];

  let i = 0;
  let inString = false;
  let bracketDepth = 0;
  let braceDepth = 0;

  while (i < cursor) {
    const ch = formula[i];

    if (inString) {
      if (ch === '"') {
        if (formula[i + 1] === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }

    if (ch === '"') {
      inString = true;
      i += 1;
      continue;
    }

    if (ch === "[") {
      bracketDepth += 1;
      i += 1;
      continue;
    }

    if (ch === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      i += 1;
      continue;
    }

    if (ch === "{") {
      braceDepth += 1;
      i += 1;
      continue;
    }

    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      i += 1;
      continue;
    }

    if (isIdentifierStart(ch)) {
      const start = i;
      i += 1;
      while (i < cursor && isIdentifierPart(formula[i])) i += 1;
      const name = formula.slice(start, i).toUpperCase();

      const next = formula[i];
      if (next === "(" && i < cursor) {
        stack.push({ kind: "function", name, argIndex: 0 });
        i += 1;
        continue;
      }

      continue;
    }

    if (ch === "(") {
      stack.push({ kind: "group" });
      i += 1;
      continue;
    }

    if (ch === ")") {
      stack.pop();
      i += 1;
      continue;
    }

    if (ch === "," || ch === ";") {
      if (bracketDepth === 0 && braceDepth === 0) {
        for (let s = stack.length - 1; s >= 0; s -= 1) {
          const frame = stack[s];
          if (frame.kind === "function") {
            frame.argIndex += 1;
            break;
          }
        }
      }
      i += 1;
      continue;
    }

    if (isWhitespace(ch)) {
      i += 1;
      continue;
    }

    i += 1;
  }

  for (let s = stack.length - 1; s >= 0; s -= 1) {
    const frame = stack[s];
    if (frame.kind === "function") {
      return { name: frame.name, argIndex: frame.argIndex };
    }
  }

  return null;
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
