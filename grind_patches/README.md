# Grind patch carrier: fastrender grid gaps + RTL fragmentation track order

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `7d640a0b6_grid_gaps_rtl_fragmentation.patch`
  - `git format-patch` output for commit `7d640a0b6f59f299b3daf9186ac59c1a74dba7a6`.

- `7d640a0b6_grid_gaps_rtl_fragmentation.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `a279ee1d464205cadd089b978e36571ffe3eab22`.

The change fixes several grid layout and fragmentation regressions:

- Minimal grid: respect `row-gap` / `column-gap`, preserve `span` when longhands override shorthands, and avoid
  stretching fixed tracks for `align-content: stretch` / `justify-content: normal|stretch`.
- Fragmentation: consume `GridTrackRanges` in global flow order when mirrored coordinates reverse logical track order,
  and select row vs column placement based on the physical fragmentation axis.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/7d640a0b6_grid_gaps_rtl_fragmentation.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/7d640a0b6_grid_gaps_rtl_fragmentation.patch
```

## Validation

The originating change was validated with:

- `cargo test -p fastrender --lib --locked`
- `cargo check -p fastrender --no-default-features --features renderer_minimal --lib --locked`
