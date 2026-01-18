# Grind patch carrier: fastrender non-macOS layout + fragmentation fixes

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `a1410110_non_macos_layout_fragmentation.patch`
  - `git format-patch` output for commit `a1410110dbc2c845b8b5aeb472b82b3c2e149bc9`.

- `a1410110_non_macos_layout_fragmentation.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `ed924cccbbc1b726c1459be227e2498fcdf7235d`.

The change fixes several layout snapshot regressions and non-macOS test failures, including:

- Minimal grid layout: resolve gaps against the container inline size and include gaps in free-space distribution.
- Fragmentation: treat mirrored grid track ranges as reversed in flow order.
- JS escapes: preserve unpaired surrogate `\\uD800..\\uDFFF` sequences verbatim.
- Test/build cleanup: fix a skipped grid fragmentation test and restore non-macOS compilation shims.
- Tooling: fix `xtask` vendoring/workspace snapshot guardrails.

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/a1410110_non_macos_layout_fragmentation.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle HEAD

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/a1410110_non_macos_layout_fragmentation.patch
```

## Validation

The originating change was validated with:

- `cargo test -p fastrender --lib --locked`
- `cargo test -p layout_engine --lib --locked`
