# Collab binder perf benchmarks

This directory contains **opt-in** `node:test` performance benchmarks for the collab binder.

They are **skipped by default** so they do not slow down CI.

## Run

Recommended (uses the repo’s TypeScript-aware node:test harness):

```bash
FORMULA_RUN_COLLAB_BINDER_PERF=1 \
NODE_OPTIONS=--expose-gc \
FORMULA_NODE_TEST_CONCURRENCY=1 \
pnpm test:node binder-perf
```

## Useful env vars

### Workload

- `PERF_CELL_UPDATES` (default `50000`)
- `PERF_BATCH_SIZE` (default `1000`)
- `PERF_COLS` (default `100`)
- `PERF_KEY_ENCODING=canonical|legacy|rxc` (default `canonical`) — used for the **Yjs → DocumentController** scenario
- `PERF_SCENARIO=yjs-to-dc|dc-to-yjs|all` (default `all`)

### Output

- `PERF_JSON=1` — emit a JSON object per scenario (easy to parse in CI)

### Optional CI-style enforcement

Disabled unless set:

- `PERF_MAX_TOTAL_MS_YJS_TO_DC` / `PERF_MAX_TOTAL_MS_DC_TO_YJS`
- `PERF_MAX_PEAK_HEAP_BYTES_YJS_TO_DC` / `PERF_MAX_PEAK_RSS_BYTES_YJS_TO_DC`
- `PERF_MAX_PEAK_HEAP_BYTES_DC_TO_YJS` / `PERF_MAX_PEAK_RSS_BYTES_DC_TO_YJS`

### Misc

- `PERF_TIMEOUT_MS` (default `600000`) — test timeout + internal wait timeouts
- `PERF_INCLUDE_GUARDS=0` — disable `canReadCell/canEditCell` hooks (enabled by default to exercise session-like binder paths)

