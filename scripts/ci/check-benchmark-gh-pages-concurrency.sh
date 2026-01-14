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
    done < <(find "$path" -maxdepth 4 -type f \( -name '*.yml' -o -name '*.yaml' \) | sort)
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
  missing_jobs="$(
    awk -v expected_group="benchmark-gh-pages-publish" '
      function indent(s) {
        match(s, /^[ ]*/);
        return RLENGTH;
      }

      function trim(s) {
        sub(/^[[:space:]]+/, "", s);
        sub(/[[:space:]]+$/, "", s);
        return s;
      }

      function strip_quotes(s) {
        s = trim(s);
        if (substr(s, 1, 1) == "\"" && substr(s, length(s), 1) == "\"") {
          return substr(s, 2, length(s) - 2);
        }
        if (substr(s, 1, 1) == "'"'"'" && substr(s, length(s), 1) == "'"'"'") {
          return substr(s, 2, length(s) - 2);
        }
        return s;
      }

      function inline_mapping_has_group(v, expected,    inner, n, parts, i, part, gv) {
        v = trim(v);
        if (substr(v, 1, 1) != "{" || substr(v, length(v), 1) != "}") {
          return 0;
        }
        inner = substr(v, 2, length(v) - 2);
        n = split(inner, parts, ",");
        for (i = 1; i <= n; i++) {
          part = trim(parts[i]);
          if (part ~ /^group[[:space:]]*:/) {
            gv = part;
            sub(/^group[[:space:]]*:/, "", gv);
            gv = strip_quotes(gv);
            if (gv == expected) return 1;
          }
        }
        return 0;
      }

      function inline_mapping_get_autopush(v,    inner, n, parts, i, part, av) {
        v = trim(v);
        if (substr(v, 1, 1) != "{" || substr(v, length(v), 1) != "}") {
          return "";
        }
        inner = substr(v, 2, length(v) - 2);
        n = split(inner, parts, ",");
        for (i = 1; i <= n; i++) {
          part = trim(parts[i]);
          if (part ~ /^auto-push[[:space:]]*:/) {
            av = part;
            sub(/^auto-push[[:space:]]*:/, "", av);
            return trim(av);
          }
        }
        return "";
      }

      function auto_push_enabled(v,    low, compact) {
        v = strip_quotes(v);
        low = tolower(v);
        compact = low;
        gsub(/[[:space:]]+/, "", compact);
        # Treat explicit "false"/0 values (including `${{ false }}`) as disabled.
        if (compact == "" ||
          compact == "false" ||
          compact == "0" ||
          compact == "${{false}}" ||
          compact == "${{0}}" ||
          compact == "${{'false'}}" ||
          compact == "${{\"false\"}}" ||
          compact == "${{'0'}}" ||
          compact == "${{\"0\"}}") {
          return 0;
        }
        return 1;
      }

      function finish_step() {
        if (step_is_benchmark && step_auto_push) {
          job_auto_push = 1;
        }
        in_step = 0;
        step_is_benchmark = 0;
        step_auto_push = 0;
        in_with = 0;
        with_indent = 0;
      }

      function finish_job() {
        finish_step();
        if (job_name != "" && job_auto_push && !(workflow_concurrency_ok || job_concurrency_ok)) {
          print job_name;
        }
        job_name = "";
        job_auto_push = 0;
        job_concurrency_ok = 0;
        in_job_concurrency = 0;
        job_concurrency_indent = 0;
        in_steps = 0;
        steps_indent = 0;
      }

      BEGIN {
        block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";

        in_block = 0;
        block_indent = 0;

        workflow_concurrency_ok = 0;
        in_workflow_concurrency = 0;
        workflow_concurrency_indent = 0;

        in_jobs = 0;
        job_name = "";
        job_auto_push = 0;
        job_concurrency_ok = 0;
        in_job_concurrency = 0;
        job_concurrency_indent = 0;

        in_steps = 0;
        steps_indent = 0;
        in_step = 0;
        step_is_benchmark = 0;
        step_auto_push = 0;
        in_with = 0;
        with_indent = 0;
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
        trimmed = line;
        sub(/^[[:space:]]*/, "", trimmed);

        # Detect YAML block scalars (e.g. `run: |`, `script: >-`) so we can skip their content.
        if (is_block) {
          in_block = 1;
          block_indent = ind;
          next;
        }

        # Workflow-level concurrency: allow serializing the entire workflow as an alternative.
        if (ind == 0 && trimmed ~ /^concurrency:[[:space:]]*/) {
          v = trimmed;
          sub(/^concurrency:[[:space:]]*/, "", v);
          v = strip_quotes(v);
          if (v != "") {
            if (v == expected_group || inline_mapping_has_group(v, expected_group)) {
              workflow_concurrency_ok = 1;
            }
          } else {
            in_workflow_concurrency = 1;
            workflow_concurrency_indent = ind;
          }
          next;
        }

        if (in_workflow_concurrency) {
          if (ind <= workflow_concurrency_indent && trimmed !~ /^$/) {
            in_workflow_concurrency = 0;
          } else if (trimmed ~ /^group:[[:space:]]*/) {
            v = trimmed;
            sub(/^group:[[:space:]]*/, "", v);
            v = strip_quotes(v);
            if (v == expected_group) {
              workflow_concurrency_ok = 1;
            }
          }
        }

        if (ind == 0 && trimmed ~ /^jobs:[[:space:]]*$/) {
          in_jobs = 1;
          next;
        }

        if (!in_jobs) {
          next;
        }

        # End the jobs map if we hit a new root key.
        if (ind == 0 && trimmed ~ /^[A-Za-z0-9_-]+:/ && trimmed !~ /^jobs:/) {
          finish_job();
          in_jobs = 0;
          next;
        }

        # Job start (under jobs:)
        if (ind == 2 && trimmed ~ /^[A-Za-z0-9_-]+:[[:space:]]*$/) {
          finish_job();
          job_name = trimmed;
          sub(/:.*/, "", job_name);
          job_name = trim(job_name);
          next;
        }

        if (job_name == "") {
          next;
        }

        # Job-level concurrency
        if (ind == 4 && trimmed ~ /^concurrency:[[:space:]]*/) {
          v = trimmed;
          sub(/^concurrency:[[:space:]]*/, "", v);
          v = strip_quotes(v);
          if (v != "") {
            if (v == expected_group || inline_mapping_has_group(v, expected_group)) {
              job_concurrency_ok = 1;
            }
            in_job_concurrency = 0;
          } else {
            in_job_concurrency = 1;
            job_concurrency_indent = ind;
          }
          next;
        }

        if (in_job_concurrency) {
          if (ind <= job_concurrency_indent && trimmed !~ /^$/ && trimmed !~ /^concurrency:/) {
            in_job_concurrency = 0;
          } else if (trimmed ~ /^group:[[:space:]]*/) {
            v = trimmed;
            sub(/^group:[[:space:]]*/, "", v);
            v = strip_quotes(v);
            if (v == expected_group) {
              job_concurrency_ok = 1;
            }
          }
        }

        # Steps list start/end
        if (ind == 4 && trimmed ~ /^steps:[[:space:]]*$/) {
          in_steps = 1;
          steps_indent = ind;
          next;
        }

        if (in_steps && ind <= steps_indent && trimmed !~ /^$/ && trimmed !~ /^steps:/) {
          finish_step();
          in_steps = 0;
        }

        if (!in_steps) {
          next;
        }

        # New step item
        if (ind == steps_indent + 2 && trimmed ~ /^-[[:space:]]/) {
          finish_step();
          in_step = 1;
          step_is_benchmark = 0;
          step_auto_push = 0;
          in_with = 0;
          with_indent = 0;
          # Continue parsing this line (it may contain inline keys).
        }

        if (!in_step) {
          next;
        }

        # Per-step parsing: detect `uses: benchmark-action/...` and `with: auto-push: ...`.
        step_line = trimmed;
        if (step_line ~ /^-[[:space:]]/) {
          sub(/^-+[[:space:]]*/, "", step_line);
        }
        step_line = trim(step_line);

        if (step_line ~ /^uses:[[:space:]]*/) {
          v = step_line;
          sub(/^uses:[[:space:]]*/, "", v);
          v = strip_quotes(v);
          if (v ~ /benchmark-action\/github-action-benchmark@/) {
            step_is_benchmark = 1;
          }
        }

        if (step_line ~ /^with:[[:space:]]*$/) {
          in_with = 1;
          with_indent = ind;
          next;
        }

        if (step_line ~ /^with:[[:space:]]*/) {
          v = step_line;
          sub(/^with:[[:space:]]*/, "", v);
          v = trim(v);
          if (v != "") {
            ap = inline_mapping_get_autopush(v);
            if (ap != "" && auto_push_enabled(ap)) {
              step_auto_push = 1;
            }
          }
        }

        if (in_with) {
          if (ind <= with_indent && trimmed !~ /^$/ && step_line !~ /^with:/) {
            in_with = 0;
          } else if (step_line ~ /^auto-push:[[:space:]]*/) {
            v = step_line;
            sub(/^auto-push:[[:space:]]*/, "", v);
            if (auto_push_enabled(v)) {
              step_auto_push = 1;
            }
          }
        }
      }

      END {
        finish_job();
      }
    ' "$wf"
  )"

  if [ -n "$missing_jobs" ]; then
    while IFS= read -r job; do
      [ -z "$job" ] && continue
      bad+=("${wf}:${job}")
    done <<<"$missing_jobs"
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
  for entry in "${bad[@]}"; do
    echo "  - ${entry}" >&2
  done
  exit 1
fi

echo "benchmark-action gh-pages concurrency: OK"
