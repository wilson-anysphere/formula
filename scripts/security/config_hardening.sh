#!/usr/bin/env bash

set -uo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

REPORT_DIR="${REPORT_DIR:-security-report}"
OUT_FILE="${OUT_FILE:-${REPORT_DIR}/config-hardening.txt}"

mkdir -p "$(dirname "$OUT_FILE")"
: >"$OUT_FILE"

fail=0

echo "== Configuration hardening checks ==" >>"$OUT_FILE"
echo "Generated: $(date -u +"%Y-%m-%dT%H:%M:%SZ")" >>"$OUT_FILE"
echo >>"$OUT_FILE"

tauri_files=()
while IFS= read -r -d '' f; do
  tauri_files+=("$f")
done < <(find . -type f -name "tauri.conf.json" -print0 2>/dev/null)

if [ ${#tauri_files[@]} -eq 0 ]; then
  echo "Tauri: skipped (no tauri.conf.json found)" >>"$OUT_FILE"
else
  echo "Tauri:" >>"$OUT_FILE"
  for f in "${tauri_files[@]}"; do
    echo "  - Checking ${f}" >>"$OUT_FILE"
    # Use Python's JSON parser to avoid jq dependency and keep rules in one place.
    python3 - "$f" >>"$OUT_FILE" 2>&1 <<'PY' || fail=1
import json
import sys
from pathlib import Path

path = Path(sys.argv[1])
data = json.loads(path.read_text(encoding="utf-8"))

def jget(obj, *keys, default=None):
    cur = obj
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur

allow_all = jget(data, "tauri", "allowlist", "all", default=False)
shell_open = jget(data, "tauri", "allowlist", "shell", "open", default=False)

csp = jget(data, "tauri", "security", "csp", default=None)
danger_disable_csp_mod = jget(data, "tauri", "security", "dangerousDisableAssetCspModification", default=False)

problems = []
if allow_all is True:
    problems.append("tauri.allowlist.all must not be true")
if shell_open is True:
    problems.append("tauri.allowlist.shell.open must not be true (avoid arbitrary URL/file opening)")
if danger_disable_csp_mod is True:
    problems.append("tauri.security.dangerousDisableAssetCspModification must not be true")

if not isinstance(csp, str) or not csp.strip():
    problems.append("tauri.security.csp must be a non-empty string")
else:
    # Fail on trivially permissive CSP patterns.
    normalized = " ".join(csp.split()).lower()
    if "default-src *" in normalized or "default-src *;" in normalized:
        problems.append("tauri.security.csp must not use `default-src *`")
    if "default-src" in normalized and "'self'" not in normalized:
        problems.append("tauri.security.csp should include `default-src 'self'` (or equivalent restriction)")

if problems:
    print("    ❌ FAIL")
    for p in problems:
        print(f"      - {p}")
    raise SystemExit(1)
else:
    print("    ✅ OK")
PY
  done
fi

echo >>"$OUT_FILE"
echo "Services (best-effort pattern checks):" >>"$OUT_FILE"

scan_roots=()
for d in apps packages crates src server api backend services; do
  if [ -d "$d" ]; then
    scan_roots+=("$d")
  fi
done

if [ ${#scan_roots[@]} -eq 0 ]; then
  echo "  - skipped (no code directories found)" >>"$OUT_FILE"
else
  # Helper: search for a pattern and fail if any matches exist.
  check_disallowed_fixed() {
    local label="$1"
    local needle="$2"

    local matches
    matches="$(grep -RInF --exclude-dir=node_modules --exclude-dir=target --exclude-dir=dist --exclude-dir=build --exclude-dir=.git "$needle" "${scan_roots[@]}" 2>/dev/null || true)"
    if [ -n "$matches" ]; then
      echo "  - ❌ ${label} (disallowed pattern: ${needle})" >>"$OUT_FILE"
      echo "$matches" | sed 's/^/    /' >>"$OUT_FILE"
      fail=1
    else
      echo "  - ✅ ${label}" >>"$OUT_FILE"
    fi
  }

  check_disallowed_fixed "Wildcard CORS response header" "Access-Control-Allow-Origin: *"
  check_disallowed_fixed "Node TLS verification disabled" "rejectUnauthorized: false"
  check_disallowed_fixed "Rust TLS verification disabled" "danger_accept_invalid_certs"
fi

exit "$fail"

