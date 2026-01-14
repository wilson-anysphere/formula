# Why does `apps/desktop/apps/desktop/` exist?

This directory intentionally mirrors a small subset of paths under
`apps/desktop/src/…` so that repo-rooted Vitest invocations keep working when
executed from within the desktop package directory.

In particular, CI (and developers) sometimes run:

```bash
pnpm -C apps/desktop exec vitest run apps/desktop/src/drawings/__tests__/selectionHandles.test.ts
```

Because `pnpm -C apps/desktop` changes the working directory to `apps/desktop/`,
that argument is interpreted relative to the package directory and resolves to:

```
apps/desktop/apps/desktop/src/…
```

The wrapper test entrypoints under `apps/desktop/apps/desktop/src/…` import and
re-export the real test suites so those commands work reliably.

Related configuration lives in:

- `apps/desktop/vite.config.ts` (includes the wrapper entrypoints and excludes
  the real suites so they don’t run twice)
- `apps/desktop/scripts/run-vitest.mjs` (normalizes paths for `pnpm vitest …`)

Do not delete this directory unless the invocation pattern above changes.
