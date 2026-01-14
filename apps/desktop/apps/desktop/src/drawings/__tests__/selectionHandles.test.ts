// Wrapper so callers can run:
//   pnpm -C apps/desktop exec vitest run apps/desktop/src/drawings/__tests__/selectionHandles.test.ts
//
// `pnpm -C apps/desktop` executes from the `apps/desktop/` directory, so paths prefixed
// with `apps/desktop/` are interpreted relative to that root. This file bridges that
// mismatch by importing the real suite.
import "../../../../../src/drawings/__tests__/selectionHandles.test.ts";

