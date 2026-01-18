# Grind patch carrier: fastrender snapshot minimal layout + fragmentation sync

This directory preserves a single fix commit for the fastrender snapshot build.

## What this is

- `f154b61cb2_snapshot_minimal_layout_fragmentation_fix.patch`
  - `git format-patch` output for commit `f154b61cb2bf93ba9281040ae1d4edf975576a09`.

- `f154b61cb2_snapshot_minimal_layout_fragmentation_fix.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to base commit `69e179da830e1e9a11227eba2d5d93425feb3146`.

## What it fixes

This patch syncs the non-macOS “minimal” layout + fragmentation harness with recent grid/style
changes, addressing snapshot regressions:

- Fragmentation: consume mirrored `GridTrackRanges` in fragmentation flow order (orthogonal/RTL),
  mapping grid line indices accordingly so pagination doesn’t panic and preserves gutter behaviour.
- Minimal grid layout: apply `grid-row-gap` / `grid-column-gap` when computing physical track ranges
  so pagination doesn’t emit gap-only fragments.
- Minimal layout sizing: plumb percentage bases through constraints/config so `%` / `calc(% + px)`
  sizes resolve against the right base.
- Non-macOS shims: export `style_minimal::computed::PositionedStyle` and update the
  `layout_engine` compile-only test harness.

## How to apply

From a clone of the fastrender snapshot branch (the commit applies cleanly on top of
`69e179da830e1e9a11227eba2d5d93425feb3146`):

### Option A (preferred): apply the bundle

```bash
cp grind_patches/f154b61cb2_snapshot_minimal_layout_fragmentation_fix.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/f154b61cb2_snapshot_minimal_layout_fragmentation_fix.patch
```

## Validation

The originating change was validated with:

- `cargo test --locked --offline`
- `cargo test -p layout_engine --test network_process_client_compile --locked --offline`

