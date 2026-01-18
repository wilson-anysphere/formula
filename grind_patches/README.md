# Grind patch carrier: fastrender snapshot grid edge cases

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `113e7229c7_snapshot_grid_edge_cases.patch`
  - `git format-patch` output for commit `113e7229c701c46a1c8513775c6c355e97bc21a4`.

- `113e7229c7_snapshot_grid_edge_cases.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `12b1ceb9497fed27db097e7d79396905bd022818`.

The change primarily fixes a native CSS Grid edge case where very large spans (e.g. `grid-row: span 65535`) caused auto-placement to fail and prevented grid fragmentation metadata from being emitted.

It also includes related snapshot build fixes and test expectation updates around physical-axis `GridTrackRanges` ordering in vertical writing modes.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/113e7229c7_snapshot_grid_edge_cases.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/113e7229c7_snapshot_grid_edge_cases.patch
```

## Validation

The originating change was validated with:

- `cargo test -p layout_engine --quiet`
- `cargo test --manifest-path crates/fastrender/Cargo.toml --lib --quiet`
- `cargo test --manifest-path crates/layout_engine_native/Cargo.toml --quiet grid_fragmentation_clamps_overflowing_line_indices`
