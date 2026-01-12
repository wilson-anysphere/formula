#!/usr/bin/env bash

set -euo pipefail

# Detect common Git merge conflict markers:
#   <<<<<<<, |||||||, =======, >>>>>>>
#
# - Only match them at the start of a line (and require the `=======` marker to be exact)
#   to avoid false positives from docs/scripts that use long `=====` separators.
# - `git grep -I` avoids scanning binary blobs like committed `.wasm` files.
pattern='^(<{7}|\|{7}|={7}$|>{7})'

# `git grep` exits:
#   0 = matches found
#   1 = no matches
#   2 = error
set +e
git grep -n -E -I "$pattern"
status=$?
set -e

if [ "$status" -eq 0 ]; then
  echo "Merge conflict markers detected. Resolve conflicts and remove conflict markers before pushing." >&2
  exit 1
elif [ "$status" -eq 1 ]; then
  exit 0
else
  echo "git grep failed with status $status" >&2
  exit "$status"
fi

