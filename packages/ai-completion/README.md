# @formula/ai-completion

Tab- and cursor-completion helpers for spreadsheet formulas and plain-cell values.

## Benchmarks

There is a lightweight micro-benchmark to help catch accidental tab-completion latency regressions.

```bash
# From repo root:
node packages/ai-completion/bench/tabCompletionEngine.bench.mjs

# Or via pnpm:
pnpm bench:tab-completion

# Or scoped to the package:
pnpm -C packages/ai-completion bench:tab-completion
```

The benchmark prints p50/p95 latency for a few common scenarios (function-name completion, range completion on a large populated column, and repeated-string pattern completion).

### CI / budgets

When `CI=1` (or when `--ci` is passed), the script exits non-zero if any scenario exceeds the p95 budget (default: 100ms).

You can tweak parameters:

```bash
node packages/ai-completion/bench/tabCompletionEngine.bench.mjs --runs 200 --warmup 50 --budget-ms 100
```
