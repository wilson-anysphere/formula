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

fail=0

# `git grep` exits:
#   0 = matches found
#   1 = no matches
#   2 = error
set +e
matches="$(git grep -n "uses: dtolnay/rust-toolchain@" -- .github/workflows/*.yml 2>/dev/null)"
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

  if [ -z "$ref" ]; then
    continue
  fi

  # dtolnay/rust-toolchain tags sometimes include a leading "v". We only use numeric toolchain tags.
  ref="${ref#v}"

  if [ "$ref" != "$channel" ]; then
    echo "Rust toolchain pin mismatch:"
    echo "  rust-toolchain.toml channel = ${channel}"
    echo "  ${file}:${line} uses dtolnay/rust-toolchain@${ref}"
    echo "  Fix: update the workflow to dtolnay/rust-toolchain@${channel} (or update rust-toolchain.toml)."
    echo
    fail=1
  fi
done <<<"$matches"

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Rust toolchain pins match rust-toolchain.toml (channel=${channel})."
