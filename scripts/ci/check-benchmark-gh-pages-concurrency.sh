#!/usr/bin/env bash
set -euo pipefail

# Guardrail: workflows that publish benchmark-action results to gh-pages must
# serialize their pushes to avoid races/flaky failures.
#
# Why:
# - `benchmark-action/github-action-benchmark` pushes to the `gh-pages` branch when
#   `auto-push: true` (or a truthy expression).
# - Multiple benchmark workflows can run concurrently (scheduled runs, push-to-main,
#   manual runs), leading to non-fast-forward push failures when two jobs try to
#   update gh-pages at the same time.
#
# Fix:
# - Add job-level concurrency to the gh-pages publishing job:
#     concurrency:
#       group: benchmark-gh-pages-publish
#       cancel-in-progress: false
#
# Usage:
#   bash scripts/ci/check-benchmark-gh-pages-concurrency.sh .github/workflows

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

inputs=("$@")
if [ "${#inputs[@]}" -eq 0 ]; then
  inputs=(".github/workflows")
fi

workflows=()
for path in "${inputs[@]}"; do
  if [ -d "$path" ]; then
    while IFS= read -r file; do
      [ -z "$file" ] && continue
      workflows+=("$file")
    done < <(find "$path" -type f \( -name '*.yml' -o -name '*.yaml' \) | sort)
    continue
  fi

  if [ -f "$path" ]; then
    workflows+=("$path")
    continue
  fi

  echo "error: workflow file/directory not found: ${path}" >&2
  exit 2
done

if [ "${#workflows[@]}" -eq 0 ]; then
  echo "error: no workflow files found from inputs: ${inputs[*]}" >&2
  exit 2
fi

bad=()
for wf in "${workflows[@]}"; do
  read -r has_benchmark has_auto_push has_concurrency < <(
    awk '
      function indent(s) {
        match(s, /^[ ]*/);
        return RLENGTH;
      }

      BEGIN {
        has_benchmark = 0;
        has_auto_push = 0;
        has_concurrency = 0;

        in_block = 0;
        block_indent = 0;
        block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      }

      {
        raw = $0;
        sub(/\r$/, "", raw);
        ind = indent(raw);

        if (in_block) {
          # Blank/whitespace-only lines can appear inside block scalars with any indentation.
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

        # Ignore single-line `run:` steps (avoid matching `auto-push:` strings inside shell).
        if (!is_block && line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]+/) {
          next;
        }

        if (line ~ /benchmark-action\/github-action-benchmark@/) {
          has_benchmark = 1;
        }

        if (line ~ /^[[:space:]]*group:[[:space:]]*["\x27]?benchmark-gh-pages-publish["\x27]?[[:space:]]*$/) {
          has_concurrency = 1;
        }

        if (match(line, /^[[:space:]]*auto-push:[[:space:]]*(.*)$/, m)) {
          v = m[1];
          gsub(/^[[:space:]]+/, "", v);
          gsub(/[[:space:]]+$/, "", v);
          if (match(v, /^"([^"]*)"$/, q)) {
            v = q[1];
          } else if (match(v, /^'\''([^'\'']*)'\''$/, q)) {
            v = q[1];
          }
          low = tolower(v);
          if (!(low == "false" || low == "0")) {
            has_auto_push = 1;
          }
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }

      END {
        printf "%d %d %d\n", has_benchmark, has_auto_push, has_concurrency;
      }
    ' "$wf"
  )

  if [ "$has_benchmark" -eq 1 ] && [ "$has_auto_push" -eq 1 ] && [ "$has_concurrency" -ne 1 ]; then
    bad+=("$wf")
  fi
done

if [ "${#bad[@]}" -gt 0 ]; then
  echo "error: benchmark-action workflows that auto-push to gh-pages must serialize pushes." >&2
  echo "Add this to the gh-pages publishing job:" >&2
  echo "" >&2
  echo "  concurrency:" >&2
  echo "    group: benchmark-gh-pages-publish" >&2
  echo "    cancel-in-progress: false" >&2
  echo "" >&2
  echo "Missing in:" >&2
  for wf in "${bad[@]}"; do
    echo "  - ${wf}" >&2
  done
  exit 1
fi

echo "benchmark-action gh-pages concurrency: OK"
