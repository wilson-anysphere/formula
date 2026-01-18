# Grind patch carrier: fastrender build + fragmentation fixes (v2)

This directory exists to preserve a change for the (currently inaccessible) `wilson-anysphere/fastrender` repository.

## What this is

- `e048af7b46_restore_build_and_fragmentation_tests.patch`
  - `git format-patch` output for commit `e048af7b4688c5671f9d57dfe779171a1b71ed26`.

- `e048af7b46_restore_build_and_fragmentation_tests.bundle`
  - A `git bundle` containing **only** that single commit.
  - It is relative to `origin/main` commit `dac082b5577bd86094eadb4c539f3893e1d2dd35`.

This supersedes the earlier `a3f26b1a0` patch carrier by fixing an additional hidden regression test issue:

- Ensure the grid fragmentation track tests are at module scope (no accidental nesting).
- Make `align-content: stretch` only stretch **auto** grid tracks (so explicit `0px` rows stay zero-length).

## How to apply to fastrender

### Option A (preferred): apply the bundle

From a clone of `wilson-anysphere/fastrender`:

```bash
cp grind_patches/e048af7b46_restore_build_and_fragmentation_tests.bundle /tmp/fix.bundle

git fetch /tmp/fix.bundle main

git cherry-pick FETCH_HEAD
```

### Option B: apply the patch

```bash
git am < grind_patches/e048af7b46_restore_build_and_fragmentation_tests.patch
```
