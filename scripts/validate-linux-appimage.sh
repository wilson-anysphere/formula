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

die() {
  echo "${SCRIPT_NAME}: error: $*" >&2
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
  - target/**/release/bundle/appimage/*.AppImage
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

select_appimage() {
  local -a found=()
  while IFS= read -r -d '' file; do
    found+=("$file")
  done < <(
    discover_appimages "$REPO_ROOT/apps/desktop/src-tauri/target"
    discover_appimages "$REPO_ROOT/target"
  )

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

  if [ "${#found[@]}" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: Multiple AppImages found. Pass --appimage to choose one:" >&2
    for f in "${found[@]}"; do
      echo "  - ${f#$REPO_ROOT/}" >&2
    done
    exit 1
  fi

  echo "${found[0]}"
}

if [ -n "$APPIMAGE_PATH" ]; then
  # Resolve relative paths against the invocation directory so we can safely `cd`.
  if [[ "$APPIMAGE_PATH" != /* ]]; then
    APPIMAGE_PATH="$ORIG_PWD/$APPIMAGE_PATH"
  fi
else
  APPIMAGE_PATH="$(select_appimage)"
fi

if [ ! -f "$APPIMAGE_PATH" ]; then
  die "AppImage does not exist: $APPIMAGE_PATH"
fi

if [ ! -s "$APPIMAGE_PATH" ]; then
  die "AppImage is empty: $APPIMAGE_PATH"
fi

if [ ! -x "$APPIMAGE_PATH" ]; then
  info "AppImage is not executable; attempting chmod +x"
  chmod +x "$APPIMAGE_PATH" || die "Failed to mark AppImage as executable: $APPIMAGE_PATH"
fi

TMPDIR="$(mktemp -d)"
cleanup() {
  rm -rf "$TMPDIR"
}
trap cleanup EXIT

APPIMAGE_BASENAME="$(basename "$APPIMAGE_PATH")"

if ! ln -s "$APPIMAGE_PATH" "$TMPDIR/$APPIMAGE_BASENAME" 2>/dev/null; then
  # Fall back to a copy if symlinks are unavailable.
  cp "$APPIMAGE_PATH" "$TMPDIR/$APPIMAGE_BASENAME"
fi
chmod +x "$TMPDIR/$APPIMAGE_BASENAME" || true

EXTRACT_LOG="$TMPDIR/appimage-extract.log"
info "Extracting AppImage (no FUSE): $APPIMAGE_PATH"
(
  cd "$TMPDIR"
  if ! "./$APPIMAGE_BASENAME" --appimage-extract >"$EXTRACT_LOG" 2>&1; then
    echo "${SCRIPT_NAME}: error: AppImage extraction failed. Output:" >&2
    tail -200 "$EXTRACT_LOG" >&2 || true
    exit 1
  fi
)

APPDIR="$TMPDIR/squashfs-root"
if [ ! -d "$APPDIR" ]; then
  die "Extraction did not produce squashfs-root/ (expected at $APPDIR)"
fi

# 3) Validate expected extracted structure.
if [ ! -e "$APPDIR/AppRun" ]; then
  die "Missing expected entrypoint: squashfs-root/AppRun"
fi
if [ ! -x "$APPDIR/AppRun" ]; then
  die "AppRun is not executable: squashfs-root/AppRun"
fi

EXPECTED_BIN="$APPDIR/usr/bin/formula-desktop"
if [ ! -e "$EXPECTED_BIN" ]; then
  die "Missing expected main binary: squashfs-root/usr/bin/formula-desktop"
fi
if [ ! -s "$EXPECTED_BIN" ]; then
  die "Main binary exists but is empty: squashfs-root/usr/bin/formula-desktop"
fi
if [ ! -x "$EXPECTED_BIN" ]; then
  die "Main binary is not executable: squashfs-root/usr/bin/formula-desktop"
fi

APPLICATIONS_DIR="$APPDIR/usr/share/applications"
if [ ! -d "$APPLICATIONS_DIR" ]; then
  die "Missing applications directory: squashfs-root/usr/share/applications/"
fi

declare -a DESKTOP_FILES=()
while IFS= read -r -d '' desktop_file; do
  DESKTOP_FILES+=("$desktop_file")
done < <(find "$APPLICATIONS_DIR" -maxdepth 1 -type f -name '*.desktop' -print0 2>/dev/null || true)

if [ "${#DESKTOP_FILES[@]}" -eq 0 ]; then
  die "No .desktop files found under squashfs-root/usr/share/applications/"
fi

# 4) Validate spreadsheet (xlsx) integration exists in at least one desktop file.
XLSX_PATTERN='xlsx|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet'
HAS_XLSX_INTEGRATION=0
for desktop_file in "${DESKTOP_FILES[@]}"; do
  if grep -Eqi "$XLSX_PATTERN" "$desktop_file"; then
    HAS_XLSX_INTEGRATION=1
    break
  fi
done

if [ "$HAS_XLSX_INTEGRATION" -ne 1 ]; then
  echo "${SCRIPT_NAME}: error: No .desktop file advertised .xlsx support." >&2
  echo "${SCRIPT_NAME}: error: Expected to find substring 'xlsx' or MIME 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in:" >&2
  for desktop_file in "${DESKTOP_FILES[@]}"; do
    echo "  - ${desktop_file#$APPDIR/}" >&2
  done
  exit 1
fi

info "OK: AppImage validated successfully"
