# Grind patch carrier: fastrender non-macOS layout + fragmentation fixes

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `bae8c4c72_non_macos_layout_fragmentation.patch`
  - `git format-patch` output for commit `bae8c4c72bf9ff313e6a44fa218a911feb6202f6`.

- `bae8c4c72_non_macos_layout_fragmentation.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `ed924cccbbc1b726c1459be227e2498fcdf7235d`.

The change fixes several layout snapshot regressions and non-macOS test failures, including:

- Minimal grid layout: resolve gaps against the container inline size and include gaps in free-space distribution.
- Fragmentation: treat mirrored grid track ranges as reversed in flow order.
- JS escapes: preserve unpaired surrogate `\\uD800..\\uDFFF` sequences verbatim.
- Test/build cleanup: fix a skipped grid fragmentation test and restore non-macOS compilation shims.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/bae8c4c72_non_macos_layout_fragmentation.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/bae8c4c72_non_macos_layout_fragmentation.patch
```

## Validation

The originating change was validated with:

- `cargo test -p fastrender --lib --locked`
- `cargo test -p layout_engine --lib --locked`
