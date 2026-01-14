#!/usr/bin/env bash
set -euo pipefail

# Guard against a subtle YAML footgun:
#
# In YAML, an unquoted colon (`:`) can be interpreted as a mapping key/value separator. This means
# workflow steps like:
#
#   - name: Guard: Something
#
# are not valid YAML and will fail to parse / run in GitHub Actions. We prefer quoting any step
# name that contains a colon, e.g.:
#
#   - name: "Guard: Something"
#
# This script is a lightweight grep-based check (no YAML parser dependency).

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

fail=0

# Find workflow `name:` fields (workflow name, job name, step name) that contain an unquoted colon.
#
# Matches (BAD):
#   name: Guard: Foo
#   - name: Guard: Foo
#
# Does not match (good):
#   name: "Guard: Foo"
#   name: 'Guard: Foo'
#   - name: "Guard: Foo"
#   - name: 'Guard: Foo'
#
# Note: this is intentionally a conservative regex; it will miss some exotic YAML quoting patterns,
# but it's sufficient to prevent the common footgun that breaks workflow parsing.
#
# Only flag colons that would break YAML parsing for plain scalars: `:` followed by whitespace (or end-of-line).
# Colons that are part of a token (e.g. `node:test`) are valid YAML and should not be flagged.
pattern='^[[:space:]]*(-[[:space:]]+)?name:[[:space:]]+[^"'"'"'].*:([[:space:]]|$)'

# Use `git grep` so we don't depend on ripgrep being installed in CI images.
# `git grep` exits:
#   0 = matches found
#   1 = no matches
#   2 = error
set +e
matches="$(git grep -nE "$pattern" -- .github/workflows 2>/dev/null)"
status=$?
set -e

if [ "$status" -eq 2 ]; then
  echo "git grep failed while scanning workflow YAML" >&2
  exit 2
fi

if [ -n "$matches" ]; then
  echo "Found workflow step names containing an unquoted ':' (invalid YAML):" >&2
  echo "$matches" >&2
  echo >&2
  echo "Fix: quote the full name string, e.g.:" >&2
  echo "  - name: \"Guard: Foo\"" >&2
  fail=1
fi

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Workflow step name colon guard: OK"
