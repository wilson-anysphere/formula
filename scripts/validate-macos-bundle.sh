#!/usr/bin/env bash
#
# Validate macOS desktop release artifacts produced by Tauri.
#
# This script is intended for CI release pipelines to catch broken bundles early:
# - Missing `.dmg` artifacts
# - DMG does not contain the expected `.app`
# - Missing/incorrect Info.plist metadata (URL scheme, file associations)
# - Invalid code signing / Gatekeeper assessment when signing is enabled
# - Missing stapled notarization tickets when notarization is configured
#
# Usage:
#   ./scripts/validate-macos-bundle.sh
#   ./scripts/validate-macos-bundle.sh --dmg path/to/Formula.dmg
#
# Signing behavior:
#   - If APPLE_CERTIFICATE is set, run codesign + spctl verification.
#   - If APPLE_ID and APPLE_PASSWORD are set, also validate stapling (notarization ticket).
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

die() {
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    # GitHub Actions log annotation.
    printf '::error::validate-macos-bundle: %s\n' "$*" >&2
  else
    printf 'error: %s\n' "$*" >&2
  fi
  exit 1
}

warn() {
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    printf '::warning::validate-macos-bundle: %s\n' "$*" >&2
  else
    printf 'warn: %s\n' "$*" >&2
  fi
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

get_expected_url_scheme() {
  # Keep this in sync with apps/desktop/src-tauri/Info.plist, which is merged into the generated
  # app bundle Info.plist during packaging.
  local config_plist="$REPO_ROOT/apps/desktop/src-tauri/Info.plist"
  if [ -f "$config_plist" ]; then
    local scheme
    set +e
    scheme="$(
      python3 - "$config_plist" <<'PY'
import plistlib
import sys

with open(sys.argv[1], "rb") as f:
    data = plistlib.load(f)

schemes = []
for url_type in data.get("CFBundleURLTypes", []) or []:
    for s in url_type.get("CFBundleURLSchemes", []) or []:
        if isinstance(s, str):
            schemes.append(s)

print(schemes[0] if schemes else "")
PY
    )"
    local status=$?
    set -e
    if [ "$status" -eq 0 ] && [ -n "$scheme" ]; then
      echo "$scheme"
      return 0
    fi
  fi

  echo "formula"
}

EXPECTED_URL_SCHEME="$(get_expected_url_scheme)"

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
try:
    with open(plist_path, "rb") as f:
        data = plistlib.load(f)
except Exception as e:
    # Exit code 2 is reserved for parse failures; exit code 1 is "valid plist, but missing scheme".
    print(str(e))
    raise SystemExit(2)

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

  if [ "$status" -eq 2 ]; then
    die "failed to parse Info.plist at ${plist_path}: ${found}"
  elif [ "$status" -ne 0 ]; then
    die "Info.plist does not declare expected URL scheme '${expected_scheme}'. Found: ${found}. (Check apps/desktop/src-tauri/Info.plist)"
  fi
}

validate_plist_file_associations() {
  local plist_path="$1"
  local required_extension="$2"
  shift 2
  local optional_extensions=("$@")

  [ -f "$plist_path" ] || die "missing Info.plist at $plist_path"

  local output
  set +e
  output="$(
    python3 - "$plist_path" "$required_extension" "${optional_extensions[@]}" <<'PY'
import plistlib
import sys

plist_path = sys.argv[1]
required_ext = sys.argv[2].lower().lstrip(".")
optional_exts = [arg.lower().lstrip(".") for arg in sys.argv[3:]]

try:
    with open(plist_path, "rb") as f:
        data = plistlib.load(f)
except Exception as e:
    # Exit code 2 is reserved for parse failures; exit code 1 is "valid plist, but missing extension".
    print(str(e))
    raise SystemExit(2)

doc_types = data.get("CFBundleDocumentTypes")
if not doc_types:
    print("CFBundleDocumentTypes is missing or empty")
    raise SystemExit(1)

found_exts = set()
for doc in doc_types or []:
    if not isinstance(doc, dict):
        continue
    exts = doc.get("CFBundleTypeExtensions") or []
    if isinstance(exts, (list, tuple)):
        for ext in exts:
            if isinstance(ext, str) and ext.strip():
                found_exts.add(ext.strip().lower().lstrip("."))

if required_ext not in found_exts:
    found = ", ".join(sorted(found_exts)) if found_exts else "(none)"
    print(f"missing required extension '{required_ext}'. Found extensions: {found}")
    raise SystemExit(1)

missing_optional = [ext for ext in optional_exts if ext and ext not in found_exts]
if missing_optional:
    # Return a human-readable warning string for the caller.
    print("missing optional extensions: " + ", ".join(missing_optional))
PY
  )"
  local status=$?
  set -e

  if [ "$status" -eq 2 ]; then
    die "failed to parse Info.plist at ${plist_path}: ${output}"
  elif [ "$status" -ne 0 ]; then
    die "Info.plist is missing file association metadata for '.${required_extension}'. Details: ${output}. (Check apps/desktop/src-tauri/Info.plist and bundle.fileAssociations in apps/desktop/src-tauri/tauri.conf.json)"
  fi

  if [ -n "$output" ]; then
    warn "Info.plist file association metadata: ${output}"
  fi
}

validate_app_bundle() {
  local app_path="$1"

  [ -d "$app_path" ] || die "expected .app directory but not found: $app_path"

  local plist_path="${app_path}/Contents/Info.plist"
  [ -f "$plist_path" ] || die "missing Contents/Info.plist in app bundle: $app_path"

  validate_plist_url_scheme "$plist_path" "$EXPECTED_URL_SCHEME"
  echo "bundle: Info.plist OK (URL scheme '${EXPECTED_URL_SCHEME}')"

  validate_plist_file_associations "$plist_path" "xlsx" "xls" "csv"
  echo "bundle: Info.plist OK (file associations include .xlsx)"

  validate_codesign "$app_path"
  validate_app_notarization "$app_path"
}

CURRENT_DMG=""
CURRENT_MOUNT_DEV=""
CURRENT_MOUNT_POINT=""
CURRENT_TMP_FILES=()
CURRENT_TMP_DIRS=()
cleanup() {
  local restore_errexit=0
  case "$-" in
    *e*) restore_errexit=1 ;;
  esac

  set +e

  if [ -n "${CURRENT_MOUNT_DEV}" ] || [ -n "${CURRENT_MOUNT_POINT}" ]; then
    local dev="${CURRENT_MOUNT_DEV}"
    local mount_point="${CURRENT_MOUNT_POINT}"
    local detached=0

    if [ -n "${dev}" ]; then
      hdiutil detach "${dev}" >/dev/null 2>&1 && detached=1
      if [ "$detached" -eq 0 ]; then
        hdiutil detach -force "${dev}" >/dev/null 2>&1 && detached=1
      fi
    fi

    # Fallback: `hdiutil detach` also accepts the mount point. This is useful when the dev-entry
    # points to a slice (e.g. /dev/diskXs1) and the detach prefers the parent disk.
    if [ "$detached" -eq 0 ] && [ -n "${mount_point}" ]; then
      hdiutil detach "${mount_point}" >/dev/null 2>&1 && detached=1
      if [ "$detached" -eq 0 ]; then
        hdiutil detach -force "${mount_point}" >/dev/null 2>&1 && detached=1
      fi
    fi

    CURRENT_MOUNT_DEV=""
    CURRENT_MOUNT_POINT=""
  fi

  if [ "${#CURRENT_TMP_FILES[@]}" -gt 0 ]; then
    rm -f "${CURRENT_TMP_FILES[@]}" >/dev/null 2>&1 || true
    CURRENT_TMP_FILES=()
  fi

  if [ "${#CURRENT_TMP_DIRS[@]}" -gt 0 ]; then
    rm -rf "${CURRENT_TMP_DIRS[@]}" >/dev/null 2>&1 || true
    CURRENT_TMP_DIRS=()
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

  if [ -z "$CURRENT_MOUNT_POINT" ]; then
    die "mounted DMG but got empty mount point from hdiutil output: $dmg"
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

find_app_in_dir() {
  local search_root="$1"
  local expected_bundle="$2"

  if [ -d "${search_root}/${expected_bundle}" ]; then
    echo "${search_root}/${expected_bundle}"
    return 0
  fi

  local apps=()
  while IFS= read -r -d '' path; do
    apps+=("$path")
  done < <(find "$search_root" -maxdepth 3 -type d -name "*.app" -print0 2>/dev/null || true)

  if [ "${#apps[@]}" -eq 1 ]; then
    warn "expected ${expected_bundle} but found ${apps[0]##*/}; using it"
    echo "${apps[0]}"
    return 0
  fi

  if [ "${#apps[@]}" -eq 0 ]; then
    die "no .app bundles found under ${search_root} (expected ${expected_bundle})"
  fi

  local listing
  listing="$(printf '%s\n' "${apps[@]}" | sed "s|^${search_root}/||" | sort)"
  die "expected ${expected_bundle} under ${search_root}, but found multiple .app bundles:
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

validate_app_notarization() {
  local app_path="$1"
  if [ -z "${APPLE_ID:-}" ] || [ -z "${APPLE_PASSWORD:-}" ]; then
    echo "notarization: skipping stapler validation (APPLE_ID/APPLE_PASSWORD not set)"
    return 0
  fi

  command -v xcrun >/dev/null || die "xcrun not found (required when notarization env is set)"

  echo "notarization: validating stapled ticket (app)..."
  if ! xcrun stapler validate "$app_path"; then
    die "stapler validation failed. Ensure notarization succeeded and the ticket was stapled to the app bundle."
  fi
}

validate_dmg_notarization() {
  local dmg_path="$1"
  if [ -z "${APPLE_ID:-}" ] || [ -z "${APPLE_PASSWORD:-}" ]; then
    echo "notarization: skipping stapler validation (APPLE_ID/APPLE_PASSWORD not set)"
    return 0
  fi

  command -v xcrun >/dev/null || die "xcrun not found (required when notarization env is set)"
  command -v spctl >/dev/null || die "spctl not found (required when notarization env is set)"

  echo "notarization: validating stapled ticket (dmg)..."
  if ! xcrun stapler validate "$dmg_path"; then
    die "stapler validation failed. Ensure notarization succeeded and the ticket was stapled to the DMG."
  fi

  echo "notarization: Gatekeeper assessment (dmg)..."
  if ! spctl -a -vv --type open "$dmg_path"; then
    die "Gatekeeper (spctl) rejected the DMG. This often indicates missing/invalid notarization."
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

  validate_app_bundle "$app_path"
  validate_dmg_notarization "$dmg"

  # Cleanly detach between DMGs to avoid piling up mounted images. The EXIT trap
  # is a safety net for failures.
  cleanup
}

validate_app_tarball() {
  local archive="$1"
  [ -f "$archive" ] || die "app archive not found: $archive"

  local extract_dir
  extract_dir="$(mktemp -d -t formula-app-archive)"
  CURRENT_TMP_DIRS+=("$extract_dir")

  echo "bundle: validating updater archive: $archive"
  if ! tar -xzf "$archive" -C "$extract_dir"; then
    die "failed to extract app archive: $archive"
  fi

  local app_path
  app_path="$(find_app_in_dir "$extract_dir" "$EXPECTED_APP_BUNDLE")"
  echo "bundle: extracted app: $app_path"

  validate_app_bundle "$app_path"

  # Clean up extracted bundles between archives to keep disk usage low. The EXIT trap is a safety net.
  cleanup
}

dedupe_lines() {
  # Read newline-delimited paths on stdin and print unique paths, preserving first-seen order.
  python3 -c $'import sys\nseen=set()\nfor line in sys.stdin.read().splitlines():\n    if not line:\n        continue\n    if line in seen:\n        continue\n    seen.add(line)\n    print(line)\n'
}

main() {
  local dmgs=()
  local app_tars=()

  if [ -n "$DMG_OVERRIDE" ]; then
    dmgs+=("$DMG_OVERRIDE")
  else
    local roots=()

    # Respect `CARGO_TARGET_DIR` when set (common in CI caching setups). Cargo interprets relative
    # paths relative to the build working directory (repo root in CI).
    if [ -n "${CARGO_TARGET_DIR:-}" ]; then
      local cargo_target_dir="${CARGO_TARGET_DIR}"
      case "$cargo_target_dir" in
        /*) ;;
        *) cargo_target_dir="$REPO_ROOT/$cargo_target_dir" ;;
      esac
      roots+=("$cargo_target_dir")
    fi

    roots+=(
      "$REPO_ROOT/apps/desktop/src-tauri/target"
      "$REPO_ROOT/apps/desktop/target"
      "$REPO_ROOT/target"
    )

    local nullglob_was_set=0
    if shopt -q nullglob; then
      nullglob_was_set=1
    fi
    shopt -s nullglob

    local root
    for root in "${roots[@]}"; do
      [ -d "$root" ] || continue

      # Fast path: use globs against the expected bundle output directories.
      dmgs+=("$root/release/bundle/dmg/"*.dmg)
      dmgs+=("$root"/*/release/bundle/dmg/*.dmg)
      app_tars+=("$root/release/bundle/macos/"*.app.tar.gz)
      app_tars+=("$root"/*/release/bundle/macos/*.app.tar.gz)
    done

    if [ "$nullglob_was_set" -eq 0 ]; then
      shopt -u nullglob
    fi

    # Fallback: traverse target roots only when the expected globs produced nothing (layout changed).
    if [ "${#dmgs[@]}" -eq 0 ] || [ "${#app_tars[@]}" -eq 0 ]; then
      for root in "${roots[@]}"; do
        [ -d "$root" ] || continue

        if [ "${#dmgs[@]}" -eq 0 ]; then
          while IFS= read -r -d '' path; do
            dmgs+=("$path")
          done < <(find "$root" -type f -path "*/release/bundle/dmg/*.dmg" -print0 2>/dev/null || true)
        fi

        if [ "${#app_tars[@]}" -eq 0 ]; then
          while IFS= read -r -d '' path; do
            app_tars+=("$path")
          done < <(find "$root" -type f -path "*/release/bundle/macos/*.app.tar.gz" -print0 2>/dev/null || true)
        fi
      done
    fi

    if [ "${#dmgs[@]}" -gt 1 ]; then
      local deduped_dmgs
      deduped_dmgs="$(printf '%s\n' "${dmgs[@]}" | dedupe_lines)"
      dmgs=()
      while IFS= read -r line; do
        [ -n "$line" ] && dmgs+=("$line")
      done <<<"$deduped_dmgs"
    fi

    if [ "${#app_tars[@]}" -gt 1 ]; then
      local deduped_app_tars
      deduped_app_tars="$(printf '%s\n' "${app_tars[@]}" | dedupe_lines)"
      app_tars=()
      while IFS= read -r line; do
        [ -n "$line" ] && app_tars+=("$line")
      done <<<"$deduped_app_tars"
    fi
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
    die "expected a DMG at apps/desktop/src-tauri/target/**/release/bundle/dmg/*.dmg, apps/desktop/target/**/release/bundle/dmg/*.dmg, or target/**/release/bundle/dmg/*.dmg (or pass --dmg)"
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

  for archive in "${app_tars[@]}"; do
    validate_app_tarball "$archive"
  done

  echo "bundle: macOS bundle validation passed."
}

main
