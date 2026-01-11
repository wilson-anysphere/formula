# @formula/spreadsheet-frontend

Shared frontend integration utilities for wiring the Formula engine (`@formula/engine`) into the canvas grid (`@formula/grid`).

This package exists so that both **web** and **desktop** frontends can reuse the same glue code:

- A1 helpers (`A1` ↔ `{row0,col0}`)
- Engine-backed cell caching + prefetching
- Grid `CellProvider` adapter with invalidation coalescing

## Entry points

This package is split into subpath exports so consumers can avoid pulling in unnecessary dependencies:

- `@formula/spreadsheet-frontend/a1` – A1 helpers (`colToName`, `toA1`, `fromA1`, `range0ToA1`)
- `@formula/spreadsheet-frontend/cache` – `EngineCellCache` (no `@formula/grid` dependency)
- `@formula/spreadsheet-frontend/grid` – `EngineGridProvider` (`CellProvider` adapter; depends on `@formula/grid`)

The root export (`@formula/spreadsheet-frontend`) re-exports everything.

## Web usage (CanvasGrid)

```ts
import { createEngineClient } from "@formula/engine";
import { EngineCellCache, EngineGridProvider } from "@formula/spreadsheet-frontend";

const engine = createEngineClient();
await engine.init();

const cache = new EngineCellCache(engine);
const provider = new EngineGridProvider({
  cache,
  rowCount: 1_000_000 + 1,
  colCount: 100 + 1,
  headers: true
});
```

Pass `provider` to `CanvasGrid`. The grid will call `provider.prefetch(range)` as the user scrolls; the provider batches and coalesces invalidations to reduce render churn.

## Notes

- `EngineGridProvider.prefetch()` matches the grid API and is **fire-and-forget**. For tests/tools that need to await fetch completion, use `prefetchAsync()`.
- `EngineGridProvider.recalculate()` is a convenience that runs an engine recalc and updates the cache + notifies subscribers.
- `EngineCellCache` is size-bounded (default `maxEntries=200_000`) to avoid unbounded growth when scrolling large sparse sheets.
