// Vitest shim for the `zod` dependency.
//
// In some CI/dev environments we run Vitest against cached/stale `node_modules` where a workspace
// package's transitive deps may be missing. `@formula/ai-tools` depends on `zod` for parameter
// validation; without it, some desktop integration suites fail at import time.
//
// This shim is intentionally minimal and only implements the subset of Zod that this repo uses:
//   - `z.string()`, `z.number()`, `z.boolean()`, `z.null()`
//   - `z.enum()`, `z.union()`, `z.array()`, `z.object()`, `z.record()`, `z.preprocess()`
//   - chained helpers: `.optional()`, `.default()`, `.min()`, `.int()`, `.positive()`, `.url()`,
//     `.refine()`, `.superRefine()`
//   - `ZodError` with `flatten()`
//
// The goal is to keep tests runnable when `zod` is missing; production builds should still rely
// on the real `zod` package.

export type ZodPath = Array<string | number>;

export type ZodIssue = {
  code: string;
  message: string;
  path: ZodPath;
};

type ParseResult<T> = { ok: true; data: T } | { ok: false; issues: ZodIssue[] };

function ok<T>(data: T): ParseResult<T> {
  return { ok: true, data };
}

function fail(path: ZodPath, message: string, code = "custom"): ParseResult<never> {
  return { ok: false, issues: [{ code, message, path }] };
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

export class ZodError extends Error {
  readonly issues: ZodIssue[];

  constructor(issues: ZodIssue[]) {
    super(issues[0]?.message ?? "Validation error");
    this.name = "ZodError";
    this.issues = issues;
  }

  flatten(): { formErrors: string[]; fieldErrors: Record<string, string[]> } {
    const formErrors: string[] = [];
    const fieldErrors: Record<string, string[]> = {};

    for (const issue of this.issues) {
      if (!issue.path || issue.path.length === 0) {
        formErrors.push(issue.message);
        continue;
      }

      const key = String(issue.path[0]);
      (fieldErrors[key] ??= []).push(issue.message);
    }

    return { formErrors, fieldErrors };
  }
}

type SuperRefineCtx = {
  addIssue: (issue: { code?: string; message: string; path?: ZodPath }) => void;
};

abstract class BaseSchema<T> {
  // Match Zod's runtime shape enough for test introspection.
  readonly _def: { typeName: string; [key: string]: any };

  protected constructor(typeName: string, extraDef: Record<string, unknown> = {}) {
    this._def = { typeName, ...extraDef };
  }

  abstract _parse(input: unknown, path: ZodPath): ParseResult<T>;

  parse(input: unknown): T {
    const res = this._parse(input, []);
    if (res.ok) return res.data;
    throw new ZodError(res.issues);
  }

  safeParse(input: unknown): { success: true; data: T } | { success: false; error: ZodError } {
    const res = this._parse(input, []);
    if (res.ok) return { success: true, data: res.data };
    return { success: false, error: new ZodError(res.issues) };
  }

  optional(): BaseSchema<T | undefined> {
    return new ZodOptional(this);
  }

  default(defaultValue: T): BaseSchema<T> {
    return new ZodDefault(this, defaultValue);
  }

  refine(check: (value: T) => boolean, message: string): BaseSchema<T> {
    return new ZodEffects(this, { kind: "refine", check, message });
  }

  superRefine(check: (value: T, ctx: SuperRefineCtx) => void): BaseSchema<T> {
    return new ZodEffects(this, { kind: "superRefine", check });
  }

  nullable(): BaseSchema<T | null> {
    return new ZodNullable(this);
  }
}

export class ZodOptional<T> extends BaseSchema<T | undefined> {
  constructor(readonly _defInnerType: BaseSchema<T>) {
    super("ZodOptional", { innerType: _defInnerType });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T | undefined> {
    if (input === undefined) return ok(undefined);
    return this._defInnerType._parse(input, path);
  }
}

export class ZodDefault<T> extends BaseSchema<T> {
  constructor(readonly _defInnerType: BaseSchema<T>, private readonly defaultValue: T) {
    super("ZodDefault", { innerType: _defInnerType });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    if (input === undefined) {
      // Validate defaults through the inner schema (mirrors real Zod behavior).
      return this._defInnerType._parse(this.defaultValue, path);
    }
    return this._defInnerType._parse(input, path);
  }
}

export class ZodNullable<T> extends BaseSchema<T | null> {
  constructor(readonly _defInnerType: BaseSchema<T>) {
    super("ZodNullable", { innerType: _defInnerType });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T | null> {
    if (input === null) return ok(null);
    return this._defInnerType._parse(input, path);
  }
}

type EffectsDef<T> =
  | { kind: "preprocess"; transform: (input: unknown) => unknown }
  | { kind: "refine"; check: (value: T) => boolean; message: string }
  | { kind: "superRefine"; check: (value: T, ctx: SuperRefineCtx) => void };

export class ZodEffects<T> extends BaseSchema<T> {
  readonly _defSchema: BaseSchema<T>;
  private readonly effect: EffectsDef<T>;

  constructor(
    schema: BaseSchema<T>,
    effect: EffectsDef<T>
  ) {
    super("ZodEffects", { schema });
    this._defSchema = schema;
    this.effect = effect;
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    if (this.effect.kind === "preprocess") {
      let next: unknown = input;
      try {
        next = this.effect.transform(input);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return fail(path, message, "custom");
      }
      return this._defSchema._parse(next, path);
    }

    const res = this._defSchema._parse(input, path);
    if (!res.ok) return res;

    if (this.effect.kind === "refine") {
      if (!this.effect.check(res.data)) {
        return fail(path, this.effect.message, "custom");
      }
      return res;
    }

    if (this.effect.kind === "superRefine") {
      const issues: ZodIssue[] = [];
      const ctx: SuperRefineCtx = {
        addIssue: (issue) => {
          const rel = issue.path ?? [];
          issues.push({
            code: issue.code ?? "custom",
            message: issue.message,
            path: [...path, ...rel],
          });
        },
      };

      try {
        this.effect.check(res.data, ctx);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        issues.push({ code: "custom", message, path });
      }

      if (issues.length) return { ok: false, issues };
      return res;
    }

    return res;
  }
}

class ZodString extends BaseSchema<string> {
  constructor(
    private readonly constraints: { minLen?: number; url?: boolean } = {}
  ) {
    super("ZodString");
  }

  min(len: number): ZodString {
    return new ZodString({ ...this.constraints, minLen: len });
  }

  url(): ZodString {
    return new ZodString({ ...this.constraints, url: true });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<string> {
    if (typeof input !== "string") {
      return fail(path, `Expected string, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }

    if (this.constraints.minLen !== undefined && input.length < this.constraints.minLen) {
      return fail(
        path,
        `String must contain at least ${this.constraints.minLen} character(s)`,
        "too_small"
      );
    }

    if (this.constraints.url) {
      try {
        // `new URL` is available in Node + browsers.
        // eslint-disable-next-line no-new
        new URL(input);
      } catch {
        return fail(path, "Invalid url", "invalid_string");
      }
    }

    return ok(input);
  }
}

class ZodNumber extends BaseSchema<number> {
  constructor(
    private readonly constraints: { int?: boolean; positive?: boolean } = {}
  ) {
    super("ZodNumber");
  }

  int(): ZodNumber {
    return new ZodNumber({ ...this.constraints, int: true });
  }

  positive(): ZodNumber {
    return new ZodNumber({ ...this.constraints, positive: true });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<number> {
    if (typeof input !== "number" || !Number.isFinite(input)) {
      return fail(path, `Expected number, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }

    if (this.constraints.int && !Number.isInteger(input)) {
      return fail(path, "Expected integer", "invalid_type");
    }

    if (this.constraints.positive && input <= 0) {
      return fail(path, "Number must be greater than 0", "too_small");
    }

    return ok(input);
  }
}

class ZodBoolean extends BaseSchema<boolean> {
  constructor() {
    super("ZodBoolean");
  }

  _parse(input: unknown, path: ZodPath): ParseResult<boolean> {
    if (typeof input !== "boolean") {
      return fail(path, `Expected boolean, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }
    return ok(input);
  }
}

class ZodNull extends BaseSchema<null> {
  constructor() {
    super("ZodNull");
  }

  _parse(input: unknown, path: ZodPath): ParseResult<null> {
    if (input !== null) {
      return fail(path, `Expected null, received ${input === undefined ? "undefined" : typeof input}`, "invalid_type");
    }
    return ok(null);
  }
}

class ZodEnum<T extends string> extends BaseSchema<T> {
  readonly options: readonly T[];

  constructor(values: readonly T[]) {
    super("ZodEnum");
    this.options = [...values];
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    if (typeof input !== "string") {
      return fail(path, `Expected string, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }
    if (!this.options.includes(input as T)) {
      return fail(path, `Invalid enum value. Expected ${this.options.join(" | ")}`, "invalid_enum_value");
    }
    return ok(input as T);
  }
}

class ZodUnion<T> extends BaseSchema<T> {
  constructor(private readonly schemas: BaseSchema<any>[]) {
    super("ZodUnion");
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    const issues: ZodIssue[] = [];
    for (const schema of this.schemas) {
      const res = schema._parse(input, path);
      if (res.ok) return res as ParseResult<T>;
      issues.push(...res.issues);
    }
    // If none match, report a generic union error at the current path.
    return issues.length ? { ok: false, issues } : fail(path, "Invalid input", "invalid_union");
  }
}

class ZodArray<T> extends BaseSchema<T[]> {
  constructor(
    private readonly element: BaseSchema<T>,
    private readonly constraints: { minLen?: number } = {}
  ) {
    super("ZodArray");
  }

  min(len: number): ZodArray<T> {
    return new ZodArray(this.element, { ...this.constraints, minLen: len });
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T[]> {
    if (!Array.isArray(input)) {
      return fail(path, `Expected array, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }

    if (this.constraints.minLen !== undefined && input.length < this.constraints.minLen) {
      return fail(path, `Array must contain at least ${this.constraints.minLen} element(s)`, "too_small");
    }

    const out: T[] = [];
    const issues: ZodIssue[] = [];
    for (let i = 0; i < input.length; i++) {
      const res = this.element._parse(input[i], [...path, i]);
      if (res.ok) out.push(res.data);
      else issues.push(...res.issues);
    }

    if (issues.length) return { ok: false, issues };
    return ok(out);
  }
}

export class ZodObject<TShape extends Record<string, BaseSchema<any>>> extends BaseSchema<{
  [K in keyof TShape]: any;
}> {
  constructor(readonly shape: TShape) {
    super("ZodObject");
  }

  _parse(input: unknown, path: ZodPath): ParseResult<any> {
    if (!isPlainObject(input)) {
      return fail(path, `Expected object, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }

    const out: Record<string, unknown> = {};
    const issues: ZodIssue[] = [];

    for (const key of Object.keys(this.shape)) {
      const schema = this.shape[key]!;
      const res = schema._parse((input as any)[key], [...path, key]);
      if (res.ok) {
        if (res.data !== undefined) out[key] = res.data;
      } else {
        issues.push(...res.issues);
      }
    }

    if (issues.length) return { ok: false, issues };
    return ok(out);
  }
}

class ZodRecord<V> extends BaseSchema<Record<string, V>> {
  constructor(private readonly valueSchema: BaseSchema<V>) {
    super("ZodRecord");
  }

  _parse(input: unknown, path: ZodPath): ParseResult<Record<string, V>> {
    if (!isPlainObject(input)) {
      return fail(path, `Expected object, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }

    const out: Record<string, V> = {};
    const issues: ZodIssue[] = [];
    for (const [k, v] of Object.entries(input)) {
      const res = this.valueSchema._parse(v, [...path, k]);
      if (res.ok) out[k] = res.data;
      else issues.push(...res.issues);
    }

    if (issues.length) return { ok: false, issues };
    return ok(out);
  }
}

export const z = {
  ZodIssueCode: { custom: "custom" as const },
  ZodObject,
  ZodEffects,
  ZodDefault,
  ZodOptional,
  ZodNullable,
  string: () => new ZodString(),
  number: () => new ZodNumber(),
  boolean: () => new ZodBoolean(),
  null: () => new ZodNull(),
  enum: <T extends string>(values: readonly T[]) => new ZodEnum(values),
  union: <T>(schemas: BaseSchema<any>[]) => new ZodUnion<T>(schemas),
  array: <T>(schema: BaseSchema<T>) => new ZodArray(schema),
  object: <TShape extends Record<string, BaseSchema<any>>>(shape: TShape) => new ZodObject(shape),
  record: <V>(schema: BaseSchema<V>) => new ZodRecord(schema),
  preprocess: <T>(transform: (input: unknown) => unknown, schema: BaseSchema<T>) =>
    new ZodEffects(schema, { kind: "preprocess", transform }),
};
