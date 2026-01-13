#!/usr/bin/env bash
#
# Validate macOS desktop release artifacts produced by Tauri.
#
# This script is intended for CI release pipelines to catch broken bundles early:
# - Missing `.dmg` artifacts
# - DMG does not contain the expected `.app`
# - Missing/incorrect Info.plist metadata (URL scheme)
# - Invalid code signing / Gatekeeper assessment when signing is enabled
# - Missing stapled notarization tickets when notarization is configured
#
# Usage:
#   ./scripts/validate-macos-bundle.sh
#   ./scripts/validate-macos-bundle.sh --dmg path/to/Formula.dmg
#
# Signing behavior:
#   - If APPLE_CERTIFICATE is set, run codesign + spctl verification.
#   - If APPLE_ID and APPLE_PASSWORD are set, also validate stapling.
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

die() {
  printf 'error: %s\n' "$*" >&2
  exit 1
}

warn() {
  printf 'warn: %s\n' "$*" >&2
}

usage() {
  cat >&2 <<'EOF'
validate-macos-bundle.sh

Validate macOS desktop release artifacts produced by Tauri (.dmg, optionally .app.tar.gz).

Options:
  --dmg <path>   Validate a specific DMG (skip artifact discovery)
  -h, --help     Show help

Environment:
  APPLE_CERTIFICATE  When non-empty, enable codesign + spctl verification.
  APPLE_ID + APPLE_PASSWORD
                     When both non-empty, additionally validate notarization stapling.
EOF
}

DMG_OVERRIDE=""
while [ "$#" -gt 0 ]; do
  case "$1" in
    --dmg)
      [ "$#" -ge 2 ] || die "--dmg requires a path"
      DMG_OVERRIDE="$2"
      shift 2
      ;;
    -h | --help)
      usage
      exit 0
      ;;
    *)
      usage
      die "unknown argument: $1"
      ;;
  esac
done

if [ "$(uname -s)" != "Darwin" ]; then
  die "this script must run on macOS (requires hdiutil/codesign/spctl/xcrun)"
fi

command -v python3 >/dev/null || die "python3 not found (required to parse tauri.conf.json and Info.plist)"

get_product_name() {
  local tauri_conf="$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json"
  if [ ! -f "$tauri_conf" ]; then
    echo "Formula"
    return 0
  fi

  # Prefer JSON parsing (robust) and fall back to a simple grep/sed if needed.
  local product
  set +e
  product="$(
    python3 - "$tauri_conf" <<'PY'
import json, sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    data = json.load(f)
print(data.get("productName", "") or "")
PY
  )"
  local status=$?
  set -e
  if [ "$status" -eq 0 ] && [ -n "$product" ]; then
    echo "$product"
    return 0
  fi

  # Best-effort fallback when python/json parsing fails.
  product="$(sed -nE 's/^[[:space:]]*"productName"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "$tauri_conf" | head -n 1)"
  if [ -n "$product" ]; then
    echo "$product"
    return 0
  fi

  echo "Formula"
}

APP_NAME="$(get_product_name)"
EXPECTED_APP_BUNDLE="${APP_NAME}.app"
EXPECTED_URL_SCHEME="formula"

validate_plist_url_scheme() {
  local plist_path="$1"
  local expected_scheme="$2"

  [ -f "$plist_path" ] || die "missing Info.plist at $plist_path"

  local found
  set +e
  found="$(
    python3 - "$plist_path" "$expected_scheme" <<'PY'
import plistlib
import sys

plist_path = sys.argv[1]
expected = sys.argv[2]
with open(plist_path, "rb") as f:
    data = plistlib.load(f)

schemes = []
for url_type in data.get("CFBundleURLTypes", []) or []:
    for scheme in (url_type.get("CFBundleURLSchemes", []) or []):
        if isinstance(scheme, str):
            schemes.append(scheme)

if expected in schemes:
    raise SystemExit(0)

unique = sorted(set(schemes))
print(", ".join(unique) if unique else "(none)")
raise SystemExit(1)
PY
  )"
  local status=$?
  set -e

  if [ "$status" -ne 0 ]; then
    die "Info.plist does not declare expected URL scheme '${expected_scheme}'. Found: ${found}. (Check apps/desktop/src-tauri/Info.plist)"
  fi
}

CURRENT_DMG=""
CURRENT_MOUNT_DEV=""
CURRENT_MOUNT_POINT=""
CURRENT_TMP_FILES=()
cleanup() {
  local restore_errexit=0
  case "$-" in
    *e*) restore_errexit=1 ;;
  esac

  set +e

  if [ -n "${CURRENT_MOUNT_DEV}" ]; then
    hdiutil detach "${CURRENT_MOUNT_DEV}" >/dev/null 2>&1 || hdiutil detach -force "${CURRENT_MOUNT_DEV}" >/dev/null 2>&1
    CURRENT_MOUNT_DEV=""
  fi

  CURRENT_MOUNT_POINT=""

  if [ "${#CURRENT_TMP_FILES[@]}" -gt 0 ]; then
    rm -f "${CURRENT_TMP_FILES[@]}" >/dev/null 2>&1 || true
    CURRENT_TMP_FILES=()
  fi

  if [ "$restore_errexit" -eq 1 ]; then
    set -e
  fi
}
trap cleanup EXIT

attach_dmg() {
  local dmg="$1"
  local attach_plist
  attach_plist="$(mktemp -t formula-hdiutil-attach)"
  CURRENT_TMP_FILES+=("$attach_plist")

  if ! hdiutil attach -nobrowse -readonly -plist "$dmg" >"$attach_plist"; then
    die "failed to mount DMG with hdiutil: $dmg"
  fi

  local parsed
  set +e
  parsed="$(
    python3 - "$attach_plist" <<'PY'
import plistlib
import sys

with open(sys.argv[1], "rb") as f:
    data = plistlib.load(f)

entities = data.get("system-entities", [])
for ent in entities:
    mp = ent.get("mount-point")
    if mp:
        print(ent.get("dev-entry", ""))
        print(mp)
        raise SystemExit(0)

raise SystemExit(1)
PY
  )"
  local status=$?
  set -e

  if [ "$status" -ne 0 ] || [ -z "$parsed" ]; then
    die "mounted DMG but could not determine mount point from hdiutil output: $dmg"
  fi

  CURRENT_MOUNT_DEV="$(printf '%s' "$parsed" | sed -n '1p')"
  CURRENT_MOUNT_POINT="$(printf '%s' "$parsed" | sed -n '2p')"

  if [ -z "$CURRENT_MOUNT_DEV" ] || [ -z "$CURRENT_MOUNT_POINT" ]; then
    die "mounted DMG but got empty mount metadata from hdiutil output: $dmg"
  fi
}

find_app_in_mount() {
  local mount_point="$1"
  local expected_bundle="$2"

  if [ -d "${mount_point}/${expected_bundle}" ]; then
    echo "${mount_point}/${expected_bundle}"
    return 0
  fi

  local apps=()
  while IFS= read -r -d '' path; do
    apps+=("$path")
  done < <(find "$mount_point" -maxdepth 3 -type d -name "*.app" -print0 2>/dev/null || true)

  if [ "${#apps[@]}" -eq 1 ]; then
    warn "expected ${expected_bundle} but found ${apps[0]##*/}; using it"
    echo "${apps[0]}"
    return 0
  fi

  if [ "${#apps[@]}" -eq 0 ]; then
    die "no .app bundles found in mounted DMG at ${mount_point} (expected ${expected_bundle})"
  fi

  local listing
  listing="$(printf '%s\n' "${apps[@]}" | sed "s|^${mount_point}/||" | sort)"
  die "expected ${expected_bundle} in mounted DMG at ${mount_point}, but found multiple .app bundles:
${listing}"
}

validate_codesign() {
  local app_path="$1"
  if [ -z "${APPLE_CERTIFICATE:-}" ]; then
    echo "signing: skipping codesign/spctl verification (APPLE_CERTIFICATE not set)"
    return 0
  fi

  command -v codesign >/dev/null || die "codesign not found (required when APPLE_CERTIFICATE is set)"
  command -v spctl >/dev/null || die "spctl not found (required when APPLE_CERTIFICATE is set)"

  echo "signing: verifying codesign..."
  if ! codesign --verify --deep --strict --verbose=2 "$app_path"; then
    die "codesign verification failed. Ensure the app is properly signed (and that signing inputs are configured)."
  fi

  echo "signing: assessing with Gatekeeper (spctl)..."
  if ! spctl --assess --type execute --verbose=2 "$app_path"; then
    die "Gatekeeper (spctl) assessment failed. This often indicates missing/invalid notarization or an invalid signature."
  fi
}

validate_notarization() {
  local dmg_path="$1"
  if [ -z "${APPLE_ID:-}" ] || [ -z "${APPLE_PASSWORD:-}" ]; then
    echo "notarization: skipping stapler validation (APPLE_ID/APPLE_PASSWORD not set)"
    return 0
  fi

  command -v xcrun >/dev/null || die "xcrun not found (required when notarization env is set)"

  echo "notarization: validating stapled ticket..."
  if ! xcrun stapler validate "$dmg_path"; then
    die "stapler validation failed. Ensure notarization succeeded and the ticket was stapled to the DMG."
  fi
}

validate_dmg() {
  local dmg="$1"
  [ -f "$dmg" ] || die "DMG not found: $dmg"

  CURRENT_DMG="$dmg"
  echo "bundle: validating DMG: $dmg"

  attach_dmg "$dmg"
  echo "bundle: mounted at ${CURRENT_MOUNT_POINT}"

  local app_path
  app_path="$(find_app_in_mount "$CURRENT_MOUNT_POINT" "$EXPECTED_APP_BUNDLE")"
  echo "bundle: found app: $app_path"

  local plist_path="${app_path}/Contents/Info.plist"
  [ -f "$plist_path" ] || die "missing Contents/Info.plist in app bundle: $app_path"

  validate_plist_url_scheme "$plist_path" "$EXPECTED_URL_SCHEME"
  echo "bundle: Info.plist OK (URL scheme '${EXPECTED_URL_SCHEME}')"

  validate_codesign "$app_path"
  validate_notarization "$dmg"

  # Cleanly detach between DMGs to avoid piling up mounted images. The EXIT trap
  # is a safety net for failures.
  cleanup
}

main() {
  local dmgs=()
  local app_tars=()

  if [ -n "$DMG_OVERRIDE" ]; then
    dmgs+=("$DMG_OVERRIDE")
  else
    local roots=(
      "$REPO_ROOT/apps/desktop/src-tauri/target"
      "$REPO_ROOT/target"
    )

    local root
    for root in "${roots[@]}"; do
      [ -d "$root" ] || continue

      # DMG artifacts.
      while IFS= read -r -d '' path; do
        dmgs+=("$path")
      done < <(find "$root" -type f -path "*/release/bundle/dmg/*.dmg" -print0 2>/dev/null || true)

      # Optional .app.tar.gz artifacts.
      while IFS= read -r -d '' path; do
        app_tars+=("$path")
      done < <(find "$root" -type f -path "*/release/bundle/macos/*.app.tar.gz" -print0 2>/dev/null || true)
    done
  fi

  if [ "${#dmgs[@]}" -eq 0 ]; then
    if [ -n "$DMG_OVERRIDE" ]; then
      die "no DMG provided; use --dmg <path>"
    fi

    warn "no DMG artifacts found."
    if [ "${#app_tars[@]}" -gt 0 ]; then
      warn "found .app.tar.gz artifacts (but DMG is required for validation):"
      printf '  %s\n' "${app_tars[@]}" >&2
    fi
    die "expected a DMG at apps/desktop/src-tauri/target/**/release/bundle/dmg/*.dmg or target/**/release/bundle/dmg/*.dmg (or pass --dmg)"
  fi

  if [ "${#app_tars[@]}" -gt 0 ]; then
    echo "bundle: discovered .app.tar.gz artifacts:"
    printf '  %s\n' "${app_tars[@]}"
  fi

  echo "bundle: expecting app bundle name: ${EXPECTED_APP_BUNDLE}"
  echo "bundle: expecting URL scheme: ${EXPECTED_URL_SCHEME}"

  local dmg
  for dmg in "${dmgs[@]}"; do
    validate_dmg "$dmg"
  done

  echo "bundle: macOS bundle validation passed."
}

main
