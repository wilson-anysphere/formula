# Grind patch carrier: fastrender build + fragmentation fixes

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `a3f26b1a0_restore_build_and_fragmentation_tests.patch`
  - `git format-patch` output for commit `a3f26b1a015884fee25122536a4099bdf38f52bd`.

- `a3f26b1a0_restore_build_and_fragmentation_tests.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `dac082b5577bd86094eadb4c539f3893e1d2dd35`.

The original change restores `cargo test --locked --offline` and `cargo test -p layout_engine --lib` by fixing:

- WebSocket module duplication + snapshot `check-cfg` feature shims
- FragmentationAnalyzer grid track order and gutter splitting
- Minimal layout percentage base hints and grid row/column gap support
- Flex out-of-flow placement profiling hook (guarded by a layout_engine test)
- Workspace / lockfile cleanups

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
# Extract bundle from this repo/branch.
# (If you fetched this branch, you can also just copy the file directly.)
cp grind_patches/a3f26b1a0_restore_build_and_fragmentation_tests.bundle /tmp/fix.bundle

# Fetch the commit from the bundle.
git fetch /tmp/fix.bundle main

# Apply it.
git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/a3f26b1a0_restore_build_and_fragmentation_tests.patch
```

