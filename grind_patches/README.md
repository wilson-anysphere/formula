# Grind patch carrier: fastrender grid gaps + RTL fragmentation track order (snapshot)

This directory preserves a single fix commit for the fastrender snapshot build.

## What this is

- `5cbcae0572_grid_gaps_rtl_fragmentation.patch`
  - `git format-patch` output for commit `5cbcae05722e5ba917cbe6b4ae1705631ba12ff3`.

- `5cbcae0572_grid_gaps_rtl_fragmentation.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to base commit `059be6f42dac1d133fb9fd13a8ec34f4f7e0e268`.

The change fixes several grid layout + fragmentation regressions in the snapshot build:

- Minimal grid: respect `row-gap`/`column-gap`, preserve spans when numeric longhands override
  shorthands, and avoid stretching fixed tracks for `align-content: stretch` /
  `justify-content: normal|stretch`.
- Fragmentation: consume `GridTrackRanges` in global flow order when mirrored coordinates reverse
  logical order (orthogonal + RTL), and use column vs row placement based on the physical
  fragmentation axis.
- Snapshot harness: dedupe WebSocket exports + non-macOS stubs.

## How to apply

From a clone of the fastrender snapshot branch (the commit applies cleanly on top of
`059be6f42dac1d133fb9fd13a8ec34f4f7e0e268`):

### Option A (preferred): apply the bundle

```bash
cp grind_patches/5cbcae0572_grid_gaps_rtl_fragmentation.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/5cbcae0572_grid_gaps_rtl_fragmentation.patch
```

## Validation

The originating change was validated with:

- `cargo test -p fastrender --lib --locked`
- `cargo check -p fastrender --no-default-features --features renderer_minimal --lib --locked`
