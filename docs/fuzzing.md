# Formula Engine Fuzzing

This repo includes **local-only** fuzzing harnesses (via `cargo-fuzz`) for the formula engine
parser and evaluator. These targets are intended to catch panics, stack overflows, and excessive
resource use when the engine is fed adversarial input.

## Targets

The fuzz crate lives at `./fuzz/` and currently provides:

- `fuzz_parse_formula`: fuzzes `formula_engine::parse_formula` (plus lowering + locale translation)
- `fuzz_eval_formula`: fuzzes parse → compile → `eval::Evaluator::eval_formula` using a small,
  deterministic in-memory `ValueResolver`

Seed corpora live under:

- `fuzz/corpus/fuzz_parse_formula/`
- `fuzz/corpus/fuzz_eval_formula/`

## Running locally

### 1) Install prerequisites

`cargo-fuzz` requires **nightly Rust** and the `cargo-fuzz` subcommand:

```bash
# Install cargo-fuzz (prefer the repo wrapper in agent environments)
bash scripts/cargo_agent.sh install cargo-fuzz
```

If you don't have nightly installed:

```bash
rustup toolchain install nightly
```

### 2) Run a short fuzz session

Fuzzing is intentionally **not** run in CI. Locally, prefer short, time-bounded runs:

```bash
# List available targets
cargo +nightly fuzz list

# Run for ~30 seconds each (recommended quick smoke test)
cargo +nightly fuzz run fuzz_parse_formula -- -max_total_time=30
cargo +nightly fuzz run fuzz_eval_formula  -- -max_total_time=30
```

### 3) Resource limits (recommended)

Fuzzers can allocate aggressively. When running inside constrained environments, wrap with the
repo memory limiter:

```bash
# Example: cap address space to 8GB while fuzzing
bash scripts/run_limited.sh --as 8G -- cargo +nightly fuzz run fuzz_eval_formula -- -max_total_time=30
```

Notes:

- `cargo-fuzz` builds artifacts under `fuzz/target/` and may use significant disk.
- Crashes and interesting inputs are saved under `fuzz/artifacts/<target>/`.

