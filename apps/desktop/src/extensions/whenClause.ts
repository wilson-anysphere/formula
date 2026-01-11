export type ContextKeyValue = string | number | boolean | null | undefined;
export type ContextKeyLookup = (key: string) => ContextKeyValue;

type Token =
  | { type: "lparen" }
  | { type: "rparen" }
  | { type: "and" }
  | { type: "or" }
  | { type: "not" }
  | { type: "eq" }
  | { type: "neq" }
  | { type: "identifier"; value: string }
  | { type: "string"; value: string }
  | { type: "number"; value: number }
  | { type: "boolean"; value: boolean }
  | { type: "eof" };

type Expr =
  | { type: "identifier"; name: string }
  | { type: "literal"; value: ContextKeyValue }
  | { type: "not"; expr: Expr }
  | { type: "and"; left: Expr; right: Expr }
  | { type: "or"; left: Expr; right: Expr }
  | { type: "eq" | "neq"; left: Expr; right: Expr };

function isIdentStart(ch: string): boolean {
  return /[A-Za-z_]/.test(ch);
}

function isIdentPart(ch: string): boolean {
  // Permit dot/colon/dash so keys like `cell.hasValue` or `view:foo` work.
  return /[A-Za-z0-9_.:-]/.test(ch);
}

function tokenize(input: string): Token[] {
  const src = String(input ?? "");
  const tokens: Token[] = [];
  let i = 0;

  const peek = () => src[i] ?? "";
  const advance = () => src[i++] ?? "";

  while (i < src.length) {
    const ch = peek();
    if (/\s/.test(ch)) {
      advance();
      continue;
    }

    if (ch === "(") {
      advance();
      tokens.push({ type: "lparen" });
      continue;
    }

    if (ch === ")") {
      advance();
      tokens.push({ type: "rparen" });
      continue;
    }

    if (ch === "!" && src[i + 1] === "=") {
      i += 2;
      tokens.push({ type: "neq" });
      continue;
    }

    if (ch === "=" && src[i + 1] === "=") {
      i += 2;
      tokens.push({ type: "eq" });
      continue;
    }

    if (ch === "&" && src[i + 1] === "&") {
      i += 2;
      tokens.push({ type: "and" });
      continue;
    }

    if (ch === "|" && src[i + 1] === "|") {
      i += 2;
      tokens.push({ type: "or" });
      continue;
    }

    if (ch === "!") {
      advance();
      tokens.push({ type: "not" });
      continue;
    }

    if (ch === "'" || ch === '"') {
      const quote = advance();
      let value = "";
      while (i < src.length) {
        const next = advance();
        if (next === quote) break;
        if (next === "\\" && i < src.length) {
          const escaped = advance();
          value += escaped;
          continue;
        }
        value += next;
      }
      tokens.push({ type: "string", value });
      continue;
    }

    if (/[0-9]/.test(ch)) {
      let raw = "";
      while (i < src.length && /[0-9.]/.test(peek())) raw += advance();
      const parsed = Number(raw);
      if (!Number.isFinite(parsed)) {
        throw new Error(`Invalid number literal in when clause: ${raw}`);
      }
      tokens.push({ type: "number", value: parsed });
      continue;
    }

    if (isIdentStart(ch)) {
      let name = "";
      while (i < src.length && isIdentPart(peek())) name += advance();
      const lower = name.toLowerCase();
      if (lower === "true") {
        tokens.push({ type: "boolean", value: true });
      } else if (lower === "false") {
        tokens.push({ type: "boolean", value: false });
      } else {
        tokens.push({ type: "identifier", value: name });
      }
      continue;
    }

    throw new Error(`Unexpected token in when clause at ${i}: ${JSON.stringify(ch)}`);
  }

  tokens.push({ type: "eof" });
  return tokens;
}

class Parser {
  private idx = 0;

  constructor(private readonly tokens: Token[]) {}

  private peek(): Token {
    return this.tokens[this.idx] ?? { type: "eof" };
  }

  private consume(): Token {
    return this.tokens[this.idx++] ?? { type: "eof" };
  }

  private expect(type: Token["type"]): Token {
    const tok = this.consume();
    if (tok.type !== type) {
      throw new Error(`Expected token ${type} but got ${tok.type}`);
    }
    return tok;
  }

  parse(): Expr {
    const expr = this.parseOr();
    if (this.peek().type !== "eof") {
      throw new Error(`Unexpected token after when clause expression: ${this.peek().type}`);
    }
    return expr;
  }

  private parseOr(): Expr {
    let left = this.parseAnd();
    while (this.peek().type === "or") {
      this.consume();
      const right = this.parseAnd();
      left = { type: "or", left, right };
    }
    return left;
  }

  private parseAnd(): Expr {
    let left = this.parseEquality();
    while (this.peek().type === "and") {
      this.consume();
      const right = this.parseEquality();
      left = { type: "and", left, right };
    }
    return left;
  }

  private parseEquality(): Expr {
    let left = this.parseUnary();
    while (this.peek().type === "eq" || this.peek().type === "neq") {
      const op = this.consume();
      const right = this.parseUnary();
      left = { type: op.type, left, right } as Expr;
    }
    return left;
  }

  private parseUnary(): Expr {
    if (this.peek().type === "not") {
      this.consume();
      return { type: "not", expr: this.parseUnary() };
    }
    return this.parsePrimary();
  }

  private parsePrimary(): Expr {
    const tok = this.peek();
    if (tok.type === "lparen") {
      this.consume();
      const inner = this.parseOr();
      this.expect("rparen");
      return inner;
    }
    if (tok.type === "identifier") {
      this.consume();
      return { type: "identifier", name: tok.value };
    }
    if (tok.type === "string") {
      this.consume();
      return { type: "literal", value: tok.value };
    }
    if (tok.type === "number") {
      this.consume();
      return { type: "literal", value: tok.value };
    }
    if (tok.type === "boolean") {
      this.consume();
      return { type: "literal", value: tok.value };
    }
    throw new Error(`Unexpected token in when clause expression: ${tok.type}`);
  }
}

export function parseWhenClause(expression: string): Expr {
  return new Parser(tokenize(expression)).parse();
}

function truthy(value: ContextKeyValue): boolean {
  if (value == null) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) && value !== 0;
  if (typeof value === "string") return value.length > 0;
  return true;
}

function evalExpr(expr: Expr, lookup: ContextKeyLookup): boolean | ContextKeyValue {
  switch (expr.type) {
    case "literal":
      return expr.value;
    case "identifier":
      return lookup(expr.name);
    case "not":
      return !truthy(evalExpr(expr.expr, lookup) as ContextKeyValue);
    case "and":
      return truthy(evalExpr(expr.left, lookup) as ContextKeyValue) && truthy(evalExpr(expr.right, lookup) as ContextKeyValue);
    case "or":
      return truthy(evalExpr(expr.left, lookup) as ContextKeyValue) || truthy(evalExpr(expr.right, lookup) as ContextKeyValue);
    case "eq": {
      const left = evalExpr(expr.left, lookup) as ContextKeyValue;
      const right = evalExpr(expr.right, lookup) as ContextKeyValue;
      return left === right;
    }
    case "neq": {
      const left = evalExpr(expr.left, lookup) as ContextKeyValue;
      const right = evalExpr(expr.right, lookup) as ContextKeyValue;
      return left !== right;
    }
    default:
      return false;
  }
}

export function evaluateWhenClause(expression: string | null | undefined, lookup: ContextKeyLookup): boolean {
  if (expression == null) return true;
  const src = String(expression).trim();
  if (src.length === 0) return true;
  try {
    const ast = parseWhenClause(src);
    return truthy(evalExpr(ast, lookup) as ContextKeyValue);
  } catch {
    // Invalid when clauses should fail closed: treat them as not satisfied.
    return false;
  }
}

