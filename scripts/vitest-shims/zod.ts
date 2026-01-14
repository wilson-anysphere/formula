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
    return new OptionalSchema(this);
  }

  default(defaultValue: T): BaseSchema<T> {
    return new DefaultSchema(this, defaultValue);
  }

  refine(check: (value: T) => boolean, message: string): BaseSchema<T> {
    return new RefineSchema(this, check, message);
  }

  superRefine(check: (value: T, ctx: SuperRefineCtx) => void): BaseSchema<T> {
    return new SuperRefineSchema(this, check);
  }
}

class OptionalSchema<T> extends BaseSchema<T | undefined> {
  constructor(private readonly inner: BaseSchema<T>) {
    super();
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T | undefined> {
    if (input === undefined) return ok(undefined);
    return this.inner._parse(input, path);
  }
}

class DefaultSchema<T> extends BaseSchema<T> {
  constructor(
    private readonly inner: BaseSchema<T>,
    private readonly defaultValue: T
  ) {
    super();
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    if (input === undefined) {
      // Validate defaults through the inner schema (mirrors real Zod behavior).
      return this.inner._parse(this.defaultValue, path);
    }
    return this.inner._parse(input, path);
  }
}

class RefineSchema<T> extends BaseSchema<T> {
  constructor(
    private readonly inner: BaseSchema<T>,
    private readonly check: (value: T) => boolean,
    private readonly message: string
  ) {
    super();
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    const res = this.inner._parse(input, path);
    if (!res.ok) return res;
    if (!this.check(res.data)) {
      return fail(path, this.message, "custom");
    }
    return res;
  }
}

class SuperRefineSchema<T> extends BaseSchema<T> {
  constructor(
    private readonly inner: BaseSchema<T>,
    private readonly check: (value: T, ctx: SuperRefineCtx) => void
  ) {
    super();
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    const res = this.inner._parse(input, path);
    if (!res.ok) return res;

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
      this.check(res.data, ctx);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      issues.push({ code: "custom", message, path });
    }

    if (issues.length) return { ok: false, issues };
    return res;
  }
}

class ZodString extends BaseSchema<string> {
  constructor(
    private readonly constraints: { minLen?: number; url?: boolean } = {}
  ) {
    super();
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
    super();
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
  _parse(input: unknown, path: ZodPath): ParseResult<boolean> {
    if (typeof input !== "boolean") {
      return fail(path, `Expected boolean, received ${input === null ? "null" : typeof input}`, "invalid_type");
    }
    return ok(input);
  }
}

class ZodNull extends BaseSchema<null> {
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
    super();
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
    super();
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
    super();
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

class ZodObject<TShape extends Record<string, BaseSchema<any>>> extends BaseSchema<{
  [K in keyof TShape]: any;
}> {
  constructor(private readonly shape: TShape) {
    super();
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
    super();
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

class ZodPreprocess<T> extends BaseSchema<T> {
  constructor(
    private readonly transform: (input: unknown) => unknown,
    private readonly inner: BaseSchema<T>
  ) {
    super();
  }

  _parse(input: unknown, path: ZodPath): ParseResult<T> {
    let next: unknown = input;
    try {
      next = this.transform(input);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return fail(path, message, "custom");
    }
    return this.inner._parse(next, path);
  }
}

export const z = {
  ZodIssueCode: { custom: "custom" as const },
  string: () => new ZodString(),
  number: () => new ZodNumber(),
  boolean: () => new ZodBoolean(),
  null: () => new ZodNull(),
  enum: <T extends string>(values: readonly T[]) => new ZodEnum(values),
  union: <T>(schemas: BaseSchema<any>[]) => new ZodUnion<T>(schemas),
  array: <T>(schema: BaseSchema<T>) => new ZodArray(schema),
  object: <TShape extends Record<string, BaseSchema<any>>>(shape: TShape) => new ZodObject(shape),
  record: <V>(schema: BaseSchema<V>) => new ZodRecord(schema),
  preprocess: <T>(transform: (input: unknown) => unknown, schema: BaseSchema<T>) => new ZodPreprocess(transform, schema),
};

