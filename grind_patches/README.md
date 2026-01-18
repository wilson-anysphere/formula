# Grind patch carrier: fastrender layout_engine snapshot fix

This directory exists to preserve a small change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `57ef6e9c8_layout_engine_snapshot_fix.patch`
  - `git format-patch` output for commit `57ef6e9c8d5fc58aa7aef9cd25435b9c7c2e27b4`.

- `57ef6e9c8_layout_engine_snapshot_fix.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `632767ffa76ad9699c3bb199dd305448adb32472`.

The original change keeps `cargo test -p layout_engine --locked` green (compile harness stubs + websocket imports + URL policy guard + blocked legacy parser token guard).

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
# Extract bundle from this repo/branch.
# (If you fetched this branch, you can also just copy the file directly.)
cp grind_patches/57ef6e9c8_layout_engine_snapshot_fix.bundle /tmp/fix.bundle

# Fetch the commit from the bundle.
git fetch /tmp/fix.bundle HEAD

# Apply it.
git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/57ef6e9c8_layout_engine_snapshot_fix.patch
```

