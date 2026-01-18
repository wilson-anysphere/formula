# Grind patch carrier: fastrender snapshot builds green (minimal grid + fragmentation)

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `29b80e7a4_snapshot_builds_green.patch`
  - `git format-patch` output for commit `29b80e7a4751984ffb792ee54790d6b6a8906216`.

- `29b80e7a4_snapshot_builds_green.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `1ae6c9dca0d5aefe2301c5e51656b0274d1e151d`.

The change fixes snapshot-test regressions in the minimal layout harness:

- Minimal grid: keep fixed/0px tracks stable under `align-content: stretch` and keep `GridTrackRanges` track-only (exclude gaps).
- Minimal layout: propagate percentage bases through nested formatting contexts so `%`/`calc()` resolve in vertical writing modes.
- Fragmentation: handle reversed grid track flow order (orthogonal + RTL) so gutters are preserved across fragmentainers.
- Grid placement parsing: apply numeric longhands before shorthand conflict normalization.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/29b80e7a4_snapshot_builds_green.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/29b80e7a4_snapshot_builds_green.patch
```

## Validation

The originating change was validated with:

- `cargo test -p fastrender --lib --locked --offline`
- `cargo check -p fastrender --no-default-features --features renderer_minimal --lib --locked --offline`
- `cargo test -p layout_engine --lib --locked --offline`
- `cargo test -p xtask --locked --offline`
