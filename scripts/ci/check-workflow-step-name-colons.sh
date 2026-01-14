#!/usr/bin/env bash
set -euo pipefail

# Guard against a subtle YAML footgun:
#
# In YAML, an unquoted colon (`:`) can be interpreted as a mapping key/value separator. This means
# workflow name fields like:
#
#   name: Guard: Something
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
# Also ignore occurrences inside YAML block scalars (e.g. within a `run: |` script body), since those
# are non-semantic text and can contain arbitrary strings that look like YAML keys.
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

workflows=()
while IFS= read -r file; do
  [ -z "$file" ] && continue
  case "$file" in
    *.yml|*.yaml) workflows+=("$file") ;;
  esac
done < <(git ls-files -- .github/workflows 2>/dev/null || true)

matches=""
for workflow in "${workflows[@]}"; do
  out="$(
    awk -v re="$pattern" '
      function indent(s) {
        match(s, /^[ ]*/);
        return RLENGTH;
      }

      BEGIN {
        in_block = 0;
        block_indent = 0;
        # Detect YAML block scalar headers (e.g. `run: |`, `restore-keys: >-`).
        block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      }

      {
        raw = $0;
        sub(/\r$/, "", raw);
        ind = indent(raw);

        if (in_block) {
          # Blank/whitespace-only lines can appear inside block scalars with any indentation.
          # Treat them as part of the scalar so we do not accidentally stop skipping early.
          if (raw ~ /^[[:space:]]*$/) {
            next;
          }
          if (ind > block_indent) {
            next;
          }
          in_block = 0;
        }

        # Strip YAML comments (best-effort; avoids tripping on documentation).
        line = raw;
        sub(/#.*/, "", line);

        is_block = (line ~ block_re);

        if (line ~ re) {
          printf "%s:%d:%s\n", FILENAME, NR, raw;
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$workflow"
  )"
  if [ -n "$out" ]; then
    if [ -n "$matches" ]; then
      matches+=$'\n'
    fi
    matches+="$out"
  fi
done

if [ -n "$matches" ]; then
  echo "Found workflow `name:` fields containing an unquoted ':' (invalid YAML):" >&2
  echo "$matches" >&2
  echo >&2
  echo "Fix: quote the full name string, e.g.:" >&2
  echo "  - name: \"Guard: Foo\"" >&2
  fail=1
fi

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Workflow name field colon guard: OK"
