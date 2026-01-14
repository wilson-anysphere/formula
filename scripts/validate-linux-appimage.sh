#!/usr/bin/env bash
#
# Validate a Tauri-produced Linux AppImage is structurally sane and contains
# expected desktop integration metadata.
#
# Usage:
#   ./scripts/validate-linux-appimage.sh
#   ./scripts/validate-linux-appimage.sh --appimage path/to/Formula.AppImage
#
# This script is intended for CI use. It performs a minimal extraction-based
# sanity check without requiring FUSE.
set -euo pipefail

SCRIPT_NAME="$(basename "$0")"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
ORIG_PWD="$(pwd)"
TMPDIR=""

die() {
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    # Emit a GitHub Actions error annotation when running in CI.
    echo "::error::validate-linux-appimage: $*" >&2
  else
    echo "${SCRIPT_NAME}: error: $*" >&2
  fi
  exit 1
}

info() {
  echo "${SCRIPT_NAME}: $*" >&2
}

usage() {
  cat <<EOF
Usage: ${SCRIPT_NAME} [--appimage <path>] [--help]

Validates a Tauri-produced Linux .AppImage for Formula Desktop.

If --appimage is not provided, the script searches common Tauri bundle output
locations:
  - apps/desktop/src-tauri/target/**/release/bundle/appimage/*.AppImage
  - apps/desktop/target/**/release/bundle/appimage/*.AppImage
  - target/**/release/bundle/appimage/*.AppImage

If CARGO_TARGET_DIR is set, it is searched first.
EOF
}

APPIMAGE_PATH=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --appimage)
      shift
      APPIMAGE_PATH="${1:-}"
      if [ -z "$APPIMAGE_PATH" ]; then
        die "--appimage requires a path"
      fi
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      die "Unknown argument: $1 (try --help)"
      ;;
  esac
done

cleanup() {
  if [ -n "${TMPDIR:-}" ] && [ -d "$TMPDIR" ]; then
    rm -rf "$TMPDIR"
  fi
}
trap cleanup EXIT

# Best-effort: keep the expected binary name in sync with
# apps/desktop/src-tauri/tauri.conf.json `mainBinaryName` (and the Rust `[[bin]]`).
EXPECTED_MAIN_BINARY="${FORMULA_APPIMAGE_MAIN_BINARY:-}"
if [ -z "$EXPECTED_MAIN_BINARY" ]; then
  tauri_conf_path="$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json"
  if [ -f "$tauri_conf_path" ] && command -v python3 >/dev/null 2>&1; then
    EXPECTED_MAIN_BINARY="$(
      python3 - "$tauri_conf_path" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print(conf.get("mainBinaryName", ""))
PY
    )"
  elif [ -f "$tauri_conf_path" ] && command -v node >/dev/null 2>&1; then
    EXPECTED_MAIN_BINARY="$(
      node -p 'const fs=require("fs");const conf=JSON.parse(fs.readFileSync(process.argv[1],"utf8")); conf.mainBinaryName ?? ""' "$tauri_conf_path" 2>/dev/null || true
    )"
  fi
  : "${EXPECTED_MAIN_BINARY:=formula-desktop}"
fi

discover_appimages() {
  local base="$1"
  if [ ! -d "$base" ]; then
    return 0
  fi

  # Tauri bundle output resembles:
  #   <target-dir>/<target-triple>/release/bundle/appimage/*.AppImage
  #   <target-dir>/release/bundle/appimage/*.AppImage
  find "$base" \
    -type f \
    -name '*.AppImage' \
    -path '*/release/bundle/appimage/*.AppImage' \
    -print0 2>/dev/null || true
}

find_appimages() {
  local -a roots=()
  # Respect `CARGO_TARGET_DIR` when set (common in CI builds).
  if [ -n "${CARGO_TARGET_DIR:-}" ]; then
    local cargo_target="${CARGO_TARGET_DIR}"
    if [[ "$cargo_target" != /* ]]; then
      cargo_target="$REPO_ROOT/$cargo_target"
    fi
    roots+=("$cargo_target")
  fi

  roots+=(
    "$REPO_ROOT/apps/desktop/src-tauri/target"
    "$REPO_ROOT/apps/desktop/target"
    "$REPO_ROOT/target"
  )

  local -a found=()
  local root
  for root in "${roots[@]}"; do
    while IFS= read -r -d '' file; do
      found+=("$file")
    done < <(discover_appimages "$root")
  done

  if [ "${#found[@]}" -eq 0 ]; then
    die "No AppImage found. Build one with Tauri, or pass --appimage <path>."
  fi

  # Deduplicate paths in case the same directory is searched twice.
  declare -A seen=()
  local -a unique=()
  local f
  for f in "${found[@]}"; do
    if [ -z "${seen["$f"]+x}" ]; then
      seen["$f"]=1
      unique+=("$f")
    fi
  done
  found=("${unique[@]}")

  # Deterministic ordering.
  mapfile -t found < <(printf '%s\n' "${found[@]}" | sort)

  printf '%s\0' "${found[@]}"
}

declare -a APPIMAGES=()

if [ -n "$APPIMAGE_PATH" ]; then
  # Resolve relative paths against the invocation directory so we can safely `cd`.
  if [[ "$APPIMAGE_PATH" != /* ]]; then
    APPIMAGE_PATH="$ORIG_PWD/$APPIMAGE_PATH"
  fi
  APPIMAGES+=("$APPIMAGE_PATH")
else
  while IFS= read -r -d '' file; do
    APPIMAGES+=("$file")
  done < <(find_appimages)
fi

validate_appimage() {
  local appimage_path="$1"

  if [ ! -f "$appimage_path" ]; then
    die "AppImage does not exist: $appimage_path"
  fi

  if [ ! -s "$appimage_path" ]; then
    die "AppImage is empty: $appimage_path"
  fi

  if [ ! -x "$appimage_path" ]; then
    info "AppImage is not executable; attempting chmod +x: $appimage_path"
    chmod +x "$appimage_path" || die "Failed to mark AppImage as executable: $appimage_path"
  fi

  TMPDIR="$(mktemp -d)"
  local appimage_basename
  appimage_basename="$(basename "$appimage_path")"

  if ! ln -s "$appimage_path" "$TMPDIR/$appimage_basename" 2>/dev/null; then
    # Fall back to a copy if symlinks are unavailable.
    cp "$appimage_path" "$TMPDIR/$appimage_basename"
  fi
  chmod +x "$TMPDIR/$appimage_basename" || true

  local extract_log
  extract_log="$TMPDIR/appimage-extract.log"
  info "Extracting AppImage (no FUSE): $appimage_path"
  (
    cd "$TMPDIR"
    if ! "./$appimage_basename" --appimage-extract >"$extract_log" 2>&1; then
      echo "${SCRIPT_NAME}: error: AppImage extraction failed for: $appimage_path" >&2
      echo "${SCRIPT_NAME}: error: Output (tail):" >&2
      tail -200 "$extract_log" >&2 || true
      die "AppImage extraction failed for: $appimage_path"
    fi
  )

  local appdir
  appdir="$TMPDIR/squashfs-root"
  if [ ! -d "$appdir" ]; then
    die "Extraction did not produce squashfs-root/ (expected at $appdir)"
  fi

  # 3) Validate expected extracted structure.
  if [ ! -e "$appdir/AppRun" ]; then
    die "Missing expected entrypoint: squashfs-root/AppRun"
  fi
  if [ ! -x "$appdir/AppRun" ]; then
    die "AppRun is not executable: squashfs-root/AppRun"
  fi

  local expected_bin
  expected_bin="$appdir/usr/bin/$EXPECTED_MAIN_BINARY"
  if [ ! -e "$expected_bin" ]; then
    die "Missing expected main binary: squashfs-root/usr/bin/$EXPECTED_MAIN_BINARY"
  fi
  if [ ! -s "$expected_bin" ]; then
    die "Main binary exists but is empty: squashfs-root/usr/bin/$EXPECTED_MAIN_BINARY"
  fi
  if [ ! -x "$expected_bin" ]; then
    die "Main binary is not executable: squashfs-root/usr/bin/$EXPECTED_MAIN_BINARY"
  fi

  local applications_dir
  applications_dir="$appdir/usr/share/applications"
  if [ ! -d "$applications_dir" ]; then
    die "Missing applications directory: squashfs-root/usr/share/applications/"
  fi

  declare -a desktop_files=()
  while IFS= read -r -d '' desktop_file; do
    desktop_files+=("$desktop_file")
  done < <(find "$applications_dir" -type f -name '*.desktop' -print0 2>/dev/null || true)

  if [ "${#desktop_files[@]}" -eq 0 ]; then
    die "No .desktop files found under squashfs-root/usr/share/applications/"
  fi

  # 4) Validate spreadsheet (xlsx) integration exists in at least one desktop file.
  local xlsx_pattern
  xlsx_pattern='xlsx|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet'
  local has_xlsx_integration=0
  local desktop_file
  for desktop_file in "${desktop_files[@]}"; do
    if grep -Eqi "$xlsx_pattern" "$desktop_file"; then
      has_xlsx_integration=1
      break
    fi
  done

  if [ "$has_xlsx_integration" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop file advertised .xlsx support for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: Expected to find substring 'xlsx' or MIME 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in:" >&2
    for desktop_file in "${desktop_files[@]}"; do
      echo "  - ${desktop_file#$appdir/}" >&2
    done
    exit 1
  fi

  # Cleanup this AppImage extraction dir early (otherwise only happens on EXIT).
  rm -rf "$TMPDIR"
  TMPDIR=""

  info "OK: AppImage validated successfully: $appimage_path"
}

if [ "${#APPIMAGES[@]}" -eq 0 ]; then
  die "Internal error: no AppImage paths to validate"
fi

if ! command -v unsquashfs >/dev/null 2>&1; then
  info "Note: 'unsquashfs' not found on PATH (package: squashfs-tools). AppImage extraction may fail without it."
fi

info "Validating ${#APPIMAGES[@]} AppImage(s)"
for appimage in "${APPIMAGES[@]}"; do
  validate_appimage "$appimage"
done
