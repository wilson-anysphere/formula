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
done < <(
  # Avoid traversing `target/` and `node_modules/` trees, which can be huge in CI (this script is
  # invoked from scripts/security/ci.sh after Rust clippy + Node installs).
  #
  # Note: `-not -path` filters do *not* prevent `find` from descending into directories. Use
  # `-prune` so we don't walk large build trees just to ignore their matches.
  find . \
    \( \
      -name '.git' -o \
      -name 'node_modules' -o \
      -name '.pnpm-store' -o \
      -name '.turbo' -o \
      -name '.cache' -o \
      -name '.vite' -o \
      -name 'playwright-report' -o \
      -name 'test-results' -o \
      -name 'dist' -o \
      -name 'build' -o \
      -name 'coverage' -o \
      -name 'target' -o \
      -path "./${REPORT_DIR}" \
    \) -prune -o \
    -type f -name "tauri.conf.json" -print0 2>/dev/null
)

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
#
# Capabilities live under `src-tauri/capabilities/*.json` and are associated to windows by
# window label using the capability file's `"windows": [...]` list.
window_labels = set()
windows = jget(data, "app", "windows", default=[])
if isinstance(windows, list):
    for w in windows:
        if isinstance(w, dict):
            label = w.get("label")
            if isinstance(label, str) and label.strip():
                window_labels.add(label.strip())

cap_dir = path.parent / "capabilities"
if cap_dir.is_dir():
    cap_files = list(cap_dir.glob("*.json"))
    for cap_file in cap_files:
        try:
            cap_data = json.loads(cap_file.read_text(encoding="utf-8"))
        except Exception as e:
            problems.append(f"Invalid JSON in capability file {cap_file}: {e}")
            continue

        identifier = cap_data.get("identifier")
        if not isinstance(identifier, str) or not identifier.strip():
            problems.append(f"{cap_file}: missing/invalid 'identifier' (expected non-empty string)")

        cap_windows = cap_data.get("windows")
        if not isinstance(cap_windows, list) or not cap_windows:
            problems.append(f"{cap_file}: windows must be a non-empty array of window labels")
        else:
            for win in cap_windows:
                if not isinstance(win, str) or not win.strip():
                    problems.append(f"{cap_file}: windows entries must be non-empty strings (got {win!r})")
                    continue
                if window_labels and win.strip() not in window_labels:
                    problems.append(
                        f"{cap_file}: references window label '{win}', but it was not found under app.windows[].label in {path}"
                    )

         perms = cap_data.get("permissions")
         if not isinstance(perms, list) or not perms:
             problems.append(f"{cap_file}: permissions must be a non-empty array")
         else:
             # Best-effort guardrail: avoid obviously over-broad permissions.
             for p in perms:
                 # Shell plugin access from the webview is intentionally not granted in this repo.
                 # External URL opening must go through a validated Rust command boundary instead.
                 if p in ("shell:allow-open", "shell:default"):
                     problems.append(f"{cap_file}: disallowed shell permission '{p}' (use a Rust command allowlist instead)")
                 if isinstance(p, str) and "allow-all" in p:
                     problems.append(f"{cap_file}: permission '{p}' looks overly broad (avoid *:allow-all)")
                 elif isinstance(p, dict):
                     pid = p.get("identifier")
                     if pid in ("shell:allow-open", "shell:default"):
                         problems.append(
                             f"{cap_file}: disallowed shell permission '{pid}' (use a Rust command allowlist instead)"
                         )
                     if isinstance(pid, str) and "allow-all" in pid:
                         problems.append(f"{cap_file}: permission '{pid}' looks overly broad (avoid *:allow-all)")
elif window_labels:
     # If we have a desktop app config but no capabilities directory, call it out explicitly.
     problems.append(f"Missing capabilities directory: {cap_dir} (expected Tauri v2 capability files under src-tauri/capabilities/)")

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
    matches="$(
      grep -RInF \
        --exclude-dir=node_modules \
        --exclude-dir=target \
        --exclude-dir=dist \
        --exclude-dir=build \
        --exclude-dir=coverage \
        --exclude-dir=.pnpm-store \
        --exclude-dir=.turbo \
        --exclude-dir=.cache \
        --exclude-dir=.vite \
        --exclude-dir=security-report \
        --exclude-dir=test-results \
        --exclude-dir=playwright-report \
        --exclude-dir=.git \
        "$needle" "${scan_roots[@]}" 2>/dev/null || true
    )"
    if [ -n "$matches" ]; then
      echo "  - ❌ ${label} (disallowed pattern: ${needle})" >>"$OUT_FILE"
      echo "$matches" | sed 's/^/    /' >>"$OUT_FILE"
      fail=1
    else
      echo "  - ✅ ${label}" >>"$OUT_FILE"
    fi
  }

  # Detect common patterns that set wildcard CORS at runtime (avoid false positives from comments/docs).
  check_disallowed_fixed "Wildcard CORS response header (\"Access-Control-Allow-Origin\", \"*\")" '"Access-Control-Allow-Origin", "*"'
  check_disallowed_fixed "Wildcard CORS response header ('Access-Control-Allow-Origin', '*')" "'Access-Control-Allow-Origin', '*'"
  check_disallowed_fixed "Node TLS verification disabled" "rejectUnauthorized: false"
  check_disallowed_fixed "Rust TLS verification disabled" "danger_accept_invalid_certs"
fi

exit "$fail"
