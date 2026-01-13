#!/usr/bin/env bash
set -euo pipefail

# Ensure CI/release workflows use the same pinned Rust version as rust-toolchain.toml.
#
# Rationale:
# - rust-toolchain.toml is the single "intended" Rust version for the repo.
# - GitHub Actions workflows also pin dtolnay/rust-toolchain@<version> for deterministic installs.
# - This script fails fast if the two diverge, so Rust upgrades are an explicit PR with CI signal.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

toolchain_file="rust-toolchain.toml"
if [ ! -f "$toolchain_file" ]; then
  echo "Missing ${toolchain_file} at repo root" >&2
  exit 1
fi

channel="$(
  awk -F'"' '
    $1 ~ /^[[:space:]]*channel[[:space:]]*=[[:space:]]*/ {
      print $2;
      exit;
    }
  ' "$toolchain_file"
)"

if [ -z "$channel" ]; then
  echo "Failed to parse toolchain channel from ${toolchain_file}" >&2
  exit 1
fi

if ! [[ "$channel" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
  echo "rust-toolchain.toml must pin an explicit Rust release (e.g. 1.92.0); found channel=${channel}" >&2
  exit 1
fi

fail=0

# `git grep` exits:
#   0 = matches found
#   1 = no matches
#   2 = error
set +e
# Scan all tracked workflow files (supports both `.yml` and `.yaml`).
matches="$(git grep -n "uses: dtolnay/rust-toolchain@" -- .github/workflows 2>/dev/null)"
status=$?
set -e

if [ "$status" -eq 2 ]; then
  echo "git grep failed while scanning workflow pins" >&2
  exit 2
fi

while IFS= read -r match; do
  [ -z "$match" ] && continue
  # Example match: ".github/workflows/ci.yml:23:      - uses: dtolnay/rust-toolchain@1.92.0"
  file="${match%%:*}"
  rest="${match#*:}"
  line="${rest%%:*}"
  ref="${match#*dtolnay/rust-toolchain@}"
  ref="${ref%%[[:space:]]*}"
  ref="${ref#v}"

  if [ -z "$ref" ]; then
    continue
  fi

  # dtolnay/rust-toolchain can be pinned either to:
  # - a toolchain version tag (e.g. `@1.92.0`), or
  # - an immutable commit SHA (supply-chain hardening), often with a trailing `# 1.92.0` comment.
  #
  # Normalize both forms into the underlying Rust toolchain version we expect.
  toolchain_ref=""
  if [[ "$ref" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    toolchain_ref="$ref"
  else
    if ! [[ "$ref" =~ ^[0-9a-fA-F]{40}$ ]]; then
      echo "Rust toolchain action must be pinned to a Rust version tag (X.Y.Z) or a full commit SHA:"
      echo "  rust-toolchain.toml channel = ${channel}"
      echo "  ${file}:${line} uses dtolnay/rust-toolchain@${ref}"
      echo "  Fix: use dtolnay/rust-toolchain@${channel}, or pin to a commit SHA with a trailing toolchain comment:"
      echo "    uses: dtolnay/rust-toolchain@<sha> # ${channel}"
      echo
      fail=1
      continue
    fi

    # Commit SHA pins: require an explicit semver toolchain comment so we can validate the intent
    # without reaching out to GitHub during CI.
    #
    # Example:
    #   uses: dtolnay/rust-toolchain@<sha> # 1.92.0
    comment="${match#*#}"
    if [ "$comment" != "$match" ]; then
      comment="${comment#"${comment%%[![:space:]]*}"}" # ltrim
      comment="${comment%%[[:space:]]*}"
      comment="${comment#v}"
      if [[ "$comment" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
        toolchain_ref="$comment"
      fi
    fi
  fi

  if [ -z "$toolchain_ref" ]; then
    echo "Rust toolchain pin could not be interpreted as a semver version:"
    echo "  rust-toolchain.toml channel = ${channel}"
    echo "  ${file}:${line} uses dtolnay/rust-toolchain@${ref}"
    echo "  Fix: pin to @${channel}, or add a trailing comment with the toolchain version:"
    echo "    uses: dtolnay/rust-toolchain@<sha> # ${channel}"
    echo
    fail=1
    continue
  fi

  if [ "$toolchain_ref" != "$channel" ]; then
    echo "Rust toolchain pin mismatch:"
    echo "  rust-toolchain.toml channel = ${channel}"
    echo "  ${file}:${line} uses dtolnay/rust-toolchain@${ref}"
    echo "  Fix: update the workflow to match ${channel} (or update rust-toolchain.toml)."
    echo
    fail=1
  fi
done <<<"$matches"

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Rust toolchain pins match rust-toolchain.toml (channel=${channel})."
