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

# Tauri v2 config uses `app.*` keys (and window-scoped capabilities) instead of the old v1
# `tauri.allowlist` / `tauri.security.*` sections.
#
# Keep best-effort support for v1 so this script remains reusable if other repos still have v1 configs.
allow_all = jget(data, "tauri", "allowlist", "all", default=None)
shell_open = jget(data, "tauri", "allowlist", "shell", "open", default=None)

csp = jget(data, "app", "security", "csp", default=None)
csp_key = "app.security.csp"
danger_disable_csp_mod = jget(data, "app", "security", "dangerousDisableAssetCspModification", default=None)
danger_key = "app.security.dangerousDisableAssetCspModification"

if csp is None:
    csp = jget(data, "tauri", "security", "csp", default=None)
    csp_key = "tauri.security.csp"

if danger_disable_csp_mod is None:
    danger_disable_csp_mod = jget(data, "tauri", "security", "dangerousDisableAssetCspModification", default=False)
    danger_key = "tauri.security.dangerousDisableAssetCspModification"

problems = []

# Legacy v1 allowlist checks (if present).
if allow_all is True:
    problems.append("tauri.allowlist.all must not be true")
if shell_open is True:
    problems.append("tauri.allowlist.shell.open must not be true (avoid arbitrary URL/file opening)")

# CSP checks (v2: app.security.csp; v1: tauri.security.csp).
if danger_disable_csp_mod is True:
    problems.append(f"{danger_key} must not be true")

if not isinstance(csp, str) or not csp.strip():
    problems.append(f"{csp_key} must be a non-empty string")
else:
    # Fail on trivially permissive CSP patterns.
    normalized = " ".join(csp.split()).lower()
    if "default-src *" in normalized or "default-src *;" in normalized:
        problems.append(f"{csp_key} must not use `default-src *`")
    if "default-src" in normalized and "'self'" not in normalized:
        problems.append(f"{csp_key} should include `default-src 'self'` (or equivalent restriction)")

# Capabilities checks (Tauri v2).
capabilities = set()
windows = jget(data, "app", "windows", default=[])
if isinstance(windows, list):
    for w in windows:
        if isinstance(w, dict) and isinstance(w.get("capabilities"), list):
            for cap in w["capabilities"]:
                if isinstance(cap, str) and cap.strip():
                    capabilities.add(cap.strip())

if capabilities:
    cap_dir = path.parent / "capabilities"
    if not cap_dir.is_dir():
        problems.append(f"Missing capabilities directory: {cap_dir} (referenced by app.windows[].capabilities)")
    else:
        cap_files = list(cap_dir.glob("*.json"))
        parsed_caps = {}
        for cap_file in cap_files:
            try:
                parsed = json.loads(cap_file.read_text(encoding="utf-8"))
            except Exception as e:
                problems.append(f"Invalid JSON in capability file {cap_file}: {e}")
                continue
            identifier = parsed.get("identifier")
            if isinstance(identifier, str) and identifier.strip():
                parsed_caps[identifier.strip()] = (cap_file, parsed)

        for cap in sorted(capabilities):
            if cap not in parsed_caps:
                problems.append(f"Capability '{cap}' is referenced but no matching JSON file was found under {cap_dir}/*.json")
                continue
            cap_file, cap_data = parsed_caps[cap]
            perms = cap_data.get("permissions")
            if not isinstance(perms, list) or not perms:
                problems.append(f"{cap_file}: permissions must be a non-empty array")
                continue
            # Best-effort guardrail: avoid obviously over-broad permissions.
            for p in perms:
                if isinstance(p, str) and "allow-all" in p:
                    problems.append(f"{cap_file}: permission '{p}' looks overly broad (avoid *:allow-all)")

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
