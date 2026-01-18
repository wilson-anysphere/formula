# Grind patch carrier: minimal grid gaps + percentage bases + JS escape fix

This directory exists to preserve a change for the (currently inaccessible)
`wilson-anysphere/fastrender` repository.

## What this is

- `7ba0fc57e_minimal_grid_gaps_percent_bases.patch`
  - `git format-patch` output for commit `7ba0fc57ee286c0fbd8c81dabf2ba9ef66b256cd`.

- `7ba0fc57e_minimal_grid_gaps_percent_bases.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `09f5965e2313033679e3f5599424fc2e25a23c66`.

The change fixes several layout and snapshot test regressions:

- Minimal grid layout: respect `row-gap` / `column-gap` when computing track ranges, free space, and
  item placement (including `FragmentNode.grid_tracks` ordering used by fragmentation tests).
- Minimal layout: propagate inline/block percentage bases through constraints/formatting contexts so
  `%` sizing is stable in vertical writing modes on non-macOS builders.
- JS escapes: preserve unpaired surrogate `\uXXXX` sequences instead of rewriting them.
- Cleanup: remove duplicate exports/statics, add missing imports, and drop obsolete selector matcher
  parity tests.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender` at `09f5965e2313033679e3f5599424fc2e25a23c66` (or
newer):

```bash
cp grind_patches/7ba0fc57e_minimal_grid_gaps_percent_bases.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle main

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/7ba0fc57e_minimal_grid_gaps_percent_bases.patch
```

## Validation

The originating change was validated with:

- `timeout -k 10 300 cargo test -p fastrender --lib`

