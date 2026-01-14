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
Usage: ${SCRIPT_NAME} [--appimage <path>] [--all] [--exec-check] [--exec-timeout <secs>] [--help]

Validates a Tauri-produced Linux .AppImage for Formula Desktop.

If --appimage is not provided, the script searches common Tauri bundle output
locations:
  - apps/desktop/src-tauri/target/**/release/bundle/appimage/*.AppImage
  - apps/desktop/target/**/release/bundle/appimage/*.AppImage
  - target/**/release/bundle/appimage/*.AppImage

If CARGO_TARGET_DIR is set, it is searched first.

Environment:
  FORMULA_APPIMAGE_MAIN_BINARY
    Override the expected main binary name inside the AppImage AppDir (defaults to
    tauri.conf.json mainBinaryName when available).
  FORMULA_VALIDATE_ALL_APPIMAGES=1
    When auto-discovering, validate all matching AppImages instead of selecting
    the most recently modified one.
  FORMULA_VALIDATE_APPIMAGE_EXEC_CHECK=1
    Additionally run a lightweight "can execute" check by invoking the extracted
    AppRun with --startup-bench (headless-friendly, exits quickly).
  FORMULA_VALIDATE_APPIMAGE_EXEC_TIMEOUT_SECS
    Timeout (seconds) for the exec check (default: 20). Requires the timeout
    command to enforce.
EOF
}

APPIMAGE_PATH=""
AUTO_DISCOVERED=1
VALIDATE_ALL=0
EXEC_CHECK_ENABLED=0
EXEC_CHECK_TIMEOUT_SECS="${FORMULA_VALIDATE_APPIMAGE_EXEC_TIMEOUT_SECS:-20}"

is_truthy() {
  # Treat common "false" values as disabled; treat any other non-empty string as enabled.
  local v="${1:-}"
  v="$(printf '%s' "$v" | tr '[:upper:]' '[:lower:]')"
  case "$v" in
    "" | 0 | false | no | n | off) return 1 ;;
    *) return 0 ;;
  esac
}

if is_truthy "${FORMULA_VALIDATE_ALL_APPIMAGES:-}"; then
  VALIDATE_ALL=1
fi
if is_truthy "${FORMULA_VALIDATE_APPIMAGE_EXEC_CHECK:-}"; then
  EXEC_CHECK_ENABLED=1
fi

while [[ $# -gt 0 ]]; do
  case "$1" in
    --appimage)
      shift
      APPIMAGE_PATH="${1:-}"
      if [ -z "$APPIMAGE_PATH" ]; then
        die "--appimage requires a path"
      fi
      AUTO_DISCOVERED=0
      shift
      ;;
    --all)
      VALIDATE_ALL=1
      shift
      ;;
    --exec-check)
      EXEC_CHECK_ENABLED=1
      shift
      ;;
    --exec-timeout)
      shift
      EXEC_CHECK_TIMEOUT_SECS="${1:-}"
      if [ -z "$EXEC_CHECK_TIMEOUT_SECS" ]; then
        die "--exec-timeout requires a value (seconds)"
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

# Expected desktop app version (from tauri.conf.json).
EXPECTED_VERSION=""
TAURI_CONF_PATH="$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json"
if [ -f "$TAURI_CONF_PATH" ]; then
  if command -v python3 >/dev/null 2>&1; then
    EXPECTED_VERSION="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print((conf.get("version") or "").strip())
PY
    )"
  fi
  if [ -z "$EXPECTED_VERSION" ] && command -v node >/dev/null 2>&1; then
    EXPECTED_VERSION="$(
      node -p 'const fs=require("fs");const conf=JSON.parse(fs.readFileSync(process.argv[1],"utf8")); String(conf.version ?? "").trim()' "$TAURI_CONF_PATH" 2>/dev/null || true
    )"
  fi
  if [ -z "$EXPECTED_VERSION" ]; then
    # Best-effort fallback when python/json parsing isn't available.
    EXPECTED_VERSION="$(sed -nE 's/^[[:space:]]*"version"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "$TAURI_CONF_PATH" | head -n 1)"
  fi
fi
if [ -z "$EXPECTED_VERSION" ]; then
  die "Unable to determine expected desktop version from $TAURI_CONF_PATH"
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
    # Caller is responsible for error handling so failures propagate correctly
    # even when `find_appimages` is used via process substitution.
    return 0
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

if [ "${#APPIMAGES[@]}" -eq 0 ]; then
  die "No AppImage found. Build one with Tauri, or pass --appimage <path>."
fi

expected_file_arch_substring() {
  # `file` prints e.g.:
  #   ELF 64-bit LSB executable, x86-64, ...
  #   ELF 64-bit LSB executable, ARM aarch64, ...
  local arch
  arch="$(uname -m)"
  case "$arch" in
    x86_64) echo "x86-64" ;;
    aarch64) echo "aarch64" ;;
    armv7l) echo "ARM" ;;
    *) echo "" ;;
  esac
}

# When auto-discovering, try to ignore stale/cross-arch AppImages so CI doesn't
# trip over cached artifacts.
if [ "$AUTO_DISCOVERED" -eq 1 ] && [ "${#APPIMAGES[@]}" -gt 1 ] && command -v file >/dev/null 2>&1; then
  expected_arch_substring="$(expected_file_arch_substring)"
  if [ -n "$expected_arch_substring" ]; then
    declare -a filtered=()
    for appimage in "${APPIMAGES[@]}"; do
      file_out="$(file -b "$appimage" 2>/dev/null || true)"
      if grep -qiF "$expected_arch_substring" <<<"$file_out"; then
        filtered+=("$appimage")
      else
        info "Skipping AppImage with mismatched architecture: $appimage (file: $file_out)"
      fi
    done
    if [ "${#filtered[@]}" -gt 0 ]; then
      APPIMAGES=("${filtered[@]}")
    fi
  fi
fi

# If multiple AppImages remain, default to validating the most recently modified
# (usually the one produced by the current build). Allow opting into validating
# all discovered AppImages via FORMULA_VALIDATE_ALL_APPIMAGES=1.
if [ "$AUTO_DISCOVERED" -eq 1 ] && [ "${#APPIMAGES[@]}" -gt 1 ] && [ "$VALIDATE_ALL" -eq 0 ]; then
  info "Multiple AppImages found; selecting the most recently modified. Use --all or set FORMULA_VALIDATE_ALL_APPIMAGES=1 to validate all."
  newest="$(
    for f in "${APPIMAGES[@]}"; do
      ts="$(stat -c '%Y' "$f" 2>/dev/null || stat -f '%m' "$f" 2>/dev/null || echo 0)"
      printf '%s\t%s\n' "$ts" "$f"
    done | sort -nr | head -n 1 | cut -f2-
  )"
  if [ -z "$newest" ]; then
    die "Failed to select an AppImage from discovered candidates"
  fi
  APPIMAGES=("$newest")
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

  # Validate OSS/compliance artifacts are shipped inside the image.
  # For Linux packages/AppImage we standardize on `usr/share/doc/<package>/`.
  local doc_dir
  doc_dir="$appdir/usr/share/doc/$EXPECTED_MAIN_BINARY"
  for filename in LICENSE NOTICE; do
    if [ ! -f "$doc_dir/$filename" ]; then
      die "Missing compliance file in AppImage: squashfs-root/usr/share/doc/$EXPECTED_MAIN_BINARY/$filename"
    fi
  done

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

  # 4) Validate the bundle advertises spreadsheet file associations via desktop metadata.
  #
  # On Linux, file associations are driven by the `MimeType=` field in the `.desktop`
  # entry. At a minimum, require the `.desktop` file advertises xlsx integration by
  # including:
  #  - the xlsx MIME type, OR
  #  - some MIME token containing the substring `xlsx` (e.g. application/x-xlsx).
  local required_xlsx_mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  local spreadsheet_mime_regex
  spreadsheet_mime_regex='xlsx|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|application/vnd\.ms-excel|application/vnd\.ms-excel\.sheet\.macroEnabled\.12|application/vnd\.ms-excel\.sheet\.binary\.macroEnabled\.12|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.template|application/vnd\.ms-excel\.template\.macroEnabled\.12|application/vnd\.ms-excel\.addin\.macroEnabled\.12|text/csv'

  local has_any_mimetype=0
  local has_spreadsheet_mime=0
  local has_xlsx_mime=0
  local has_xlsx_integration=0
  local bad_exec_count=0

  local desktop_file
  for desktop_file in "${desktop_files[@]}"; do
    local mime_line
    mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    if [ -z "$mime_line" ]; then
      continue
    fi

    has_any_mimetype=1

    local mime_value
    mime_value="$(printf '%s' "$mime_line" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//")"

    if printf '%s' "$mime_value" | grep -Fqi "$required_xlsx_mime"; then
      has_xlsx_mime=1
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Fqi "xlsx"; then
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Eqi "$spreadsheet_mime_regex"; then
      has_spreadsheet_mime=1

      # File associations require a placeholder token (%U/%u/%F/%f) in Exec= so the
      # OS passes the opened file path/URL into the app.
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [ -z "$exec_line" ]; then
        echo "${SCRIPT_NAME}: error: ${desktop_file#$appdir/} is missing an Exec= entry (required for file associations)" >&2
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        echo "${SCRIPT_NAME}: error: ${desktop_file#$appdir/} Exec= does not include a file/URL placeholder (%U/%u/%F/%f): ${exec_line}" >&2
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi
  done

  if [ "$has_any_mimetype" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop file contained a MimeType= entry for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: This usually means Linux file associations were not included in the bundle." >&2
    echo "${SCRIPT_NAME}: error: Check apps/desktop/src-tauri/tauri.conf.json → bundle.fileAssociations." >&2
    echo "${SCRIPT_NAME}: error: .desktop files inspected:" >&2
    for desktop_file in "${desktop_files[@]}"; do
      echo "  - ${desktop_file#$appdir/}" >&2
    done
    die "No .desktop file contained a MimeType= entry for AppImage: $appimage_path"
  fi

  if [ "$has_xlsx_integration" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop MimeType= entry advertised xlsx support for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: Expected MimeType= to include substring 'xlsx' or MIME '${required_xlsx_mime}'." >&2
    echo "${SCRIPT_NAME}: error: MimeType entries found:" >&2
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#$appdir/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    die "No .desktop file advertised .xlsx support for AppImage: $appimage_path"
  fi

  if [ "$has_spreadsheet_mime" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop MimeType= entry advertised spreadsheet (xlsx) support for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: Expected MimeType= to include '${required_xlsx_mime}' (xlsx) or another spreadsheet MIME type." >&2
    echo "${SCRIPT_NAME}: error: MimeType entries found:" >&2
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#$appdir/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        # Print the raw line(s) for debugging.
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    die "No .desktop file advertised spreadsheet MIME types for AppImage: $appimage_path"
  fi

  if [ "$bad_exec_count" -ne 0 ]; then
    die "One or more .desktop entries had invalid Exec= lines for file association handling"
  fi

  if [ "$has_xlsx_mime" -ne 1 ]; then
    info "warn: No .desktop file explicitly listed xlsx MIME '${required_xlsx_mime}'. Spreadsheet MIME types were present, but .xlsx double-click integration may be incomplete."
  fi

  # 5) Validate version metadata matches tauri.conf.json. Prefer the AppImage-specific
  # X-AppImage-Version desktop entry key; otherwise accept a semver-looking Version=
  # field. If no application version marker is present, fall back to validating the
  # artifact filename includes the expected version (best-effort).
  local found_version_marker=0
  local found_value=""
  for desktop_file in "${desktop_files[@]}"; do
    found_value="$(grep -E '^X-AppImage-Version=' "$desktop_file" | head -n 1 | sed 's/^X-AppImage-Version=//' | tr -d '\r' || true)"
    if [ -n "$found_value" ]; then
      found_version_marker=1
      if [ "$found_value" != "$EXPECTED_VERSION" ]; then
        die "AppImage version mismatch (X-AppImage-Version) in ${desktop_file#$appdir/}: expected ${EXPECTED_VERSION}, found ${found_value}"
      fi
    fi
  done

  if [ "$found_version_marker" -eq 0 ]; then
    for desktop_file in "${desktop_files[@]}"; do
      found_value="$(grep -E '^Version=' "$desktop_file" | head -n 1 | sed 's/^Version=//' | tr -d '\r' || true)"
      # Note: Desktop Entry "Version" is often the spec version (commonly 1.0), not
      # the application version. Only treat semver-like values as an app version.
      if [[ -n "$found_value" && "$found_value" =~ ^[0-9]+\.[0-9]+\.[0-9]+([.-].*)?$ ]]; then
        found_version_marker=1
        if [ "$found_value" != "$EXPECTED_VERSION" ]; then
          die "AppImage version mismatch (Version=) in ${desktop_file#$appdir/}: expected ${EXPECTED_VERSION}, found ${found_value}"
        fi
      fi
    done
  fi

  if [ "$found_version_marker" -eq 0 ]; then
    local appimage_filename
    appimage_filename="$(basename "$appimage_path")"
    if [[ "$appimage_filename" != *"$EXPECTED_VERSION"* ]]; then
      die "AppImage did not expose X-AppImage-Version/Version in its desktop entry, and filename did not contain expected version ${EXPECTED_VERSION}: ${appimage_filename}"
    fi
  fi

  # Optional (recommended): ensure the extracted AppRun can execute without FUSE by
  # running in a mode that exits quickly. This is disabled by default because it
  # may require a working display (Xvfb) and WebKit/GTK runtime dependencies.
  if [ "$EXEC_CHECK_ENABLED" -ne 0 ]; then
    local timeout_secs="${EXEC_CHECK_TIMEOUT_SECS}"
    if ! [[ "$timeout_secs" =~ ^[0-9]+$ ]] || [ "$timeout_secs" -lt 1 ]; then
      die "Invalid exec timeout: ${timeout_secs} (expected integer >= 1)"
    fi

    info "Exec check: running extracted AppRun --startup-bench (timeout ${timeout_secs}s)"
    local exec_log
    exec_log="$TMPDIR/apprun-exec.log"

    # Run under a virtual display when available/needed.
    local -a runner=()
    if [ -x "$REPO_ROOT/scripts/xvfb-run-safe.sh" ] && { [ -z "${DISPLAY:-}" ] || [ -n "${CI:-}" ]; }; then
      runner=("$REPO_ROOT/scripts/xvfb-run-safe.sh")
    fi

    local status=0
    set +e
    if command -v timeout >/dev/null 2>&1; then
      "${runner[@]}" timeout "${timeout_secs}s" bash -lc "cd \"$appdir\" && ./AppRun --startup-bench" >"$exec_log" 2>&1
      status=$?
    else
      info "Exec check: 'timeout' not found; running without an enforced timeout"
      "${runner[@]}" bash -lc "cd \"$appdir\" && ./AppRun --startup-bench" >"$exec_log" 2>&1
      status=$?
    fi
    set -e

    if [ "$status" -ne 0 ]; then
      echo "${SCRIPT_NAME}: error: AppRun --startup-bench failed (exit ${status}) for AppImage: $appimage_path" >&2
      tail -200 "$exec_log" >&2 || true
      die "Extracted AppRun did not execute successfully for AppImage: $appimage_path"
    fi
  fi

  # Cleanup this AppImage extraction dir early (otherwise only happens on EXIT).
  rm -rf "$TMPDIR"
  TMPDIR=""

  info "OK: AppImage validated successfully: $appimage_path"
}

if [ "${#APPIMAGES[@]}" -eq 0 ]; then
  die "No AppImage paths to validate after discovery. Build an AppImage or pass --appimage <path>."
fi

if ! command -v unsquashfs >/dev/null 2>&1; then
  info "Note: 'unsquashfs' not found on PATH (package: squashfs-tools). AppImage extraction may fail without it."
fi

info "Validating ${#APPIMAGES[@]} AppImage(s)"
for appimage in "${APPIMAGES[@]}"; do
  validate_appimage "$appimage"
done
