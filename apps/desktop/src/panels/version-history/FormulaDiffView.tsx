import React, { useMemo } from "react";

import { diffFormula } from "../../versioning/index.js";

type DiffOpType = "equal" | "insert" | "delete";

type FormulaToken = { type: string; value: string };

export function FormulaDiffView({
  oldFormula,
  newFormula,
  className
}: {
  oldFormula: string | null | undefined;
  newFormula: string | null | undefined;
  className?: string;
}) {
  const flat = useMemo(() => {
    const result = diffFormula(oldFormula ?? null, newFormula ?? null);
    const out: Array<{ op: DiffOpType; token: FormulaToken }> = [];
    for (const op of result.ops) {
      for (const token of op.tokens) out.push({ op: op.type, token });
    }
    return out;
  }, [oldFormula, newFormula]);

  if (flat.length === 0) {
    return <span className="branch-merge__empty">∅</span>;
  }

  return (
    <pre className={["formula-bar-highlight", "branch-merge__formula-diff", className].filter(Boolean).join(" ")}>
      {flat.map(({ op, token }, i) => {
        const nextToken = flat[i + 1]?.token ?? null;
        const kind = tokenKind(token, nextToken);
        return (
          <span
            // eslint-disable-next-line react/no-array-index-key
            key={i}
            className={["branch-merge__diff-token", `branch-merge__diff-token--${op}`].join(" ")}
            data-kind={kind}
          >
            {renderTokenText(token)}
          </span>
        );
      })}
    </pre>
  );
}

function renderTokenText(token: FormulaToken): string {
  if (token.type === "string") {
    // `diffFormula` tokenizes string literals without the surrounding quotes.
    // Re-add quotes for user-facing rendering.
    const escaped = token.value.replaceAll('"', '""');
    return `"${escaped}"`;
  }

  if (token.type === "ident") {
    // Quoted sheet names are tokenized as `ident` without the surrounding apostrophes.
    // Best-effort: re-add quotes when whitespace is present (unquoted sheet names cannot contain spaces).
    if (/\s/.test(token.value)) {
      return `'${token.value.replaceAll("'", "''")}'`;
    }
  }

  return token.value;
}

function tokenKind(token: FormulaToken, nextToken: FormulaToken | null): string {
  switch (token.type) {
    case "number":
      return "number";
    case "string":
      return "string";
    case "punct":
      return "punctuation";
    case "op":
      return "operator";
    case "ident": {
      // Heuristics for nicer syntax highlighting: try to distinguish functions and
      // A1-style cell references, otherwise treat as plain identifier.
      if (nextToken?.type === "punct" && nextToken.value === "(") return "function";

      const v = token.value;
      const isCellRef = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(v);
      if (isCellRef) return "reference";

      return "identifier";
    }
    default:
      return "identifier";
  }
}

