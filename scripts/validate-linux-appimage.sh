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

Options:
  --appimage <path>        Validate a specific .AppImage file (skip discovery).
  --all                    Validate all discovered AppImages (otherwise the newest is selected).
  --exec-check             Run a lightweight "can execute" check (AppRun --startup-bench).
  --exec-timeout <secs>    Timeout for --exec-check (default: 20).
  -h, --help               Show this help.

If --appimage is not provided, the script searches common Tauri bundle output
locations:
  - apps/desktop/src-tauri/target/release/bundle/appimage/*.AppImage
  - apps/desktop/src-tauri/target/*/release/bundle/appimage/*.AppImage
  - apps/desktop/target/release/bundle/appimage/*.AppImage
  - apps/desktop/target/*/release/bundle/appimage/*.AppImage
  - target/release/bundle/appimage/*.AppImage
  - target/*/release/bundle/appimage/*.AppImage

If CARGO_TARGET_DIR is set, it is searched first.

Environment:
  FORMULA_APPIMAGE_MAIN_BINARY
    Override the expected main binary name inside the AppImage AppDir (defaults to
    tauri.conf.json mainBinaryName when available).
  FORMULA_TAURI_CONF_PATH
    Optional path override for apps/desktop/src-tauri/tauri.conf.json (useful for local testing).
    If the path is relative, it is resolved relative to the repo root.
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

TAURI_CONF_PATH="${FORMULA_TAURI_CONF_PATH:-$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json}"
if [[ "${TAURI_CONF_PATH}" != /* ]]; then
  TAURI_CONF_PATH="${REPO_ROOT}/${TAURI_CONF_PATH}"
fi

is_truthy() {
  # Treat common "false" values as disabled; treat any other non-empty string as enabled.
  local v="${1:-}"
  # Trim whitespace and normalize case.
  v="$(printf '%s' "$v" | tr '[:upper:]' '[:lower:]' | sed -E 's/^[[:space:]]+//; s/[[:space:]]+$//')"
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
  if [ -f "$TAURI_CONF_PATH" ] && command -v python3 >/dev/null 2>&1; then
    EXPECTED_MAIN_BINARY="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print(conf.get("mainBinaryName", ""))
PY
    )"
  elif [ -f "$TAURI_CONF_PATH" ] && command -v node >/dev/null 2>&1; then
    EXPECTED_MAIN_BINARY="$(
      node -p 'const fs=require("fs");const conf=JSON.parse(fs.readFileSync(process.argv[1],"utf8")); conf.mainBinaryName ?? ""' "$TAURI_CONF_PATH" 2>/dev/null || true
    )"
  fi
  : "${EXPECTED_MAIN_BINARY:=formula-desktop}"
fi

# Expected desktop app version (from tauri.conf.json).
EXPECTED_VERSION=""
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

# Expected app identifier (reverse-DNS) from tauri.conf.json. We use it for the
# shared-mime-info definition filename under:
#   usr/share/mime/packages/<identifier>.xml
EXPECTED_IDENTIFIER=""
if [ -f "$TAURI_CONF_PATH" ]; then
  if command -v python3 >/dev/null 2>&1; then
    EXPECTED_IDENTIFIER="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print((conf.get("identifier") or "").strip())
PY
    )"
  fi
  if [ -z "$EXPECTED_IDENTIFIER" ] && command -v node >/dev/null 2>&1; then
    EXPECTED_IDENTIFIER="$(
      node -p 'const fs=require("fs");const conf=JSON.parse(fs.readFileSync(process.argv[1],"utf8")); String(conf.identifier ?? "").trim()' "$TAURI_CONF_PATH" 2>/dev/null || true
    )"
  fi
  if [ -z "$EXPECTED_IDENTIFIER" ]; then
    EXPECTED_IDENTIFIER="$(sed -nE 's/^[[:space:]]*"identifier"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "$TAURI_CONF_PATH" | head -n 1)"
  fi
fi
if [ -z "$EXPECTED_IDENTIFIER" ]; then
  die "Unable to determine expected desktop identifier from $TAURI_CONF_PATH"
fi
# The app identifier is used as a filename on Linux:
#   /usr/share/mime/packages/<identifier>.xml
if [[ "${EXPECTED_IDENTIFIER}" == */* || "${EXPECTED_IDENTIFIER}" == *\\* ]]; then
  die "Expected $TAURI_CONF_PATH identifier to be a valid filename (no '/' or '\\' path separators). Found: ${EXPECTED_IDENTIFIER}"
fi
EXPECTED_MIME_DEFINITION_BASENAME="${EXPECTED_IDENTIFIER}.xml"

# Expected deep-link URL scheme handlers advertised via desktop integration.
#
# Source of truth is apps/desktop/src-tauri/tauri.conf.json → plugins.deep-link.desktop.schemes.
# Accept either:
#   - desktop: { schemes: ["formula"] }
#   - desktop: [{ schemes: ["formula"] }, { schemes: ["formula-beta"] }]
#
# Default to "formula" if nothing is configured or the config cannot be parsed.
declare -a EXPECTED_DEEP_LINK_SCHEMES=()
if [ -f "$TAURI_CONF_PATH" ] && command -v python3 >/dev/null 2>&1; then
  while IFS= read -r line; do
    if [ -n "$line" ]; then
      EXPECTED_DEEP_LINK_SCHEMES+=("$line")
    fi
  done < <(
    python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import re
import sys

path = sys.argv[1]
try:
    with open(path, "r", encoding="utf-8") as f:
        conf = json.load(f)
except Exception:
    sys.exit(0)

plugins = conf.get("plugins") or {}
deep_link = plugins.get("deep-link") or {}
desktop = deep_link.get("desktop")
schemes = set()

def normalize(value: str) -> str:
    v = value.strip().lower()
    v = re.sub(r"[:/]+$", "", v)
    return v

def add_from_protocol(protocol):
    if not isinstance(protocol, dict):
        return
    raw = protocol.get("schemes")
    values = []
    if isinstance(raw, str):
        values = [raw]
    elif isinstance(raw, list):
        values = [v for v in raw if isinstance(v, str)]
    for v in values:
        norm = normalize(v)
        if norm:
            schemes.add(norm)

if isinstance(desktop, list):
    for protocol in desktop:
        add_from_protocol(protocol)
else:
    add_from_protocol(desktop)

for scheme in sorted(schemes):
    print(scheme)
PY
  )
fi

if [ "${#EXPECTED_DEEP_LINK_SCHEMES[@]}" -eq 0 ] && [ -f "$TAURI_CONF_PATH" ] && command -v node >/dev/null 2>&1; then
  while IFS= read -r line; do
    if [ -n "$line" ]; then
      EXPECTED_DEEP_LINK_SCHEMES+=("$line")
    fi
  done < <(
    node - "$TAURI_CONF_PATH" <<'NODE' 2>/dev/null || true
const fs = require("node:fs");
const configPath = process.argv[2];
let conf;
try {
  conf = JSON.parse(fs.readFileSync(configPath, "utf8"));
} catch {
  process.exit(0);
}

const deepLink = conf?.plugins?.["deep-link"];
const desktop = deepLink?.desktop;
const schemes = new Set();

const normalize = (value) => String(value).trim().toLowerCase().replace(/[:/]+$/, "");

const addFromProtocol = (protocol) => {
  if (!protocol || typeof protocol !== "object") return;
  const raw = protocol.schemes;
  const values = typeof raw === "string" ? [raw] : Array.isArray(raw) ? raw : [];
  for (const v of values) {
    if (typeof v !== "string") continue;
    const norm = normalize(v);
    if (norm) schemes.add(norm);
  }
};

if (Array.isArray(desktop)) {
  for (const protocol of desktop) addFromProtocol(protocol);
} else {
  addFromProtocol(desktop);
}

if (schemes.size === 0) process.exit(0);
for (const scheme of Array.from(schemes).sort()) console.log(scheme);
NODE
  )
fi
if [ "${#EXPECTED_DEEP_LINK_SCHEMES[@]}" -eq 0 ]; then
  EXPECTED_DEEP_LINK_SCHEMES=("formula")
fi

for scheme in "${EXPECTED_DEEP_LINK_SCHEMES[@]}"; do
  if [[ "$scheme" == *:* || "$scheme" == */* ]]; then
    die "Invalid deep-link scheme configured in ${TAURI_CONF_PATH} (plugins[\"deep-link\"].desktop.schemes): '${scheme}'. Expected scheme names only (no ':' or '/' characters)."
  fi
done
declare -a EXPECTED_SCHEME_MIMES=()
for scheme in "${EXPECTED_DEEP_LINK_SCHEMES[@]}"; do
  EXPECTED_SCHEME_MIMES+=("x-scheme-handler/${scheme}")
done

discover_appimages() {
  local base="$1"
  if [ ! -d "$base" ]; then
    return 0
  fi

  # Fast path: avoid traversing an entire Cargo target directory (which can be very
  # large) by checking the canonical bundle locations directly.
  #
  # Tauri bundle output resembles:
  #   <target-dir>/<target-triple>/release/bundle/appimage/*.AppImage
  #   <target-dir>/release/bundle/appimage/*.AppImage
  local nullglob_was_set=0
  if shopt -q nullglob; then
    nullglob_was_set=1
  fi
  shopt -s nullglob

  local -a matches=(
    "$base"/release/bundle/appimage/*.AppImage
    "$base"/*/release/bundle/appimage/*.AppImage
  )

  if [[ "$nullglob_was_set" -eq 0 ]]; then
    shopt -u nullglob
  fi

  if [ "${#matches[@]}" -gt 0 ]; then
    printf '%s\0' "${matches[@]}"
    return 0
  fi

  # Fallback: traverse only when the output layout is unexpected.
  find "$base" \
    -maxdepth 6 \
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
    aarch64|arm64) echo "aarch64" ;;
    armv7l) echo "ARM" ;;
    *) echo "" ;;
  esac
}

expected_arch_substring="$(expected_file_arch_substring)"

# Best-effort architecture sanity check using `file(1)`.
#
# This helps avoid trying to execute a cached AppImage for the wrong architecture, which
# otherwise fails later with a confusing "Exec format error".
if command -v file >/dev/null 2>&1 && [ -n "$expected_arch_substring" ]; then
  if [ "$AUTO_DISCOVERED" -eq 1 ]; then
    declare -a filtered=()
    declare -a mismatched=()

    for appimage in "${APPIMAGES[@]}"; do
      file_out="$(file -b "$appimage" 2>/dev/null || true)"
      if grep -qi "ELF" <<<"$file_out"; then
        if grep -qiF "$expected_arch_substring" <<<"$file_out"; then
          filtered+=("$appimage")
        else
          mismatched+=("$appimage (file: $file_out)")
          info "Skipping AppImage with mismatched architecture: $appimage (file: $file_out)"
        fi
      else
        # If the file isn't an ELF, don't apply arch filtering here; later extraction will
        # fail and provide a clearer error message. (Also keeps unit tests using a fake
        # shell-script "AppImage" working.)
        filtered+=("$appimage")
      fi
    done

    if [ "${#filtered[@]}" -eq 0 ]; then
      echo "${SCRIPT_NAME}: error: Found AppImage artifacts, but none match the host architecture (uname -m=$(uname -m))." >&2
      echo "${SCRIPT_NAME}: error: Expected to find architecture substring: ${expected_arch_substring}" >&2
      echo "${SCRIPT_NAME}: error: Mismatched AppImages:" >&2
      for line in "${mismatched[@]}"; do
        echo "  - $line" >&2
      done
      die "No AppImages matched expected architecture"
    fi

    APPIMAGES=("${filtered[@]}")
  else
    # User provided an explicit AppImage. If it's an ELF but doesn't match the host arch,
    # fail early with a clear message.
    file_out="$(file -b "${APPIMAGES[0]}" 2>/dev/null || true)"
    if grep -qi "ELF" <<<"$file_out" && ! grep -qiF "$expected_arch_substring" <<<"$file_out"; then
      die "Wrong AppImage architecture: expected '${expected_arch_substring}' for host $(uname -m), got: ${file_out}"
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

  # We ship a shared-mime-info definition for Parquet so `*.parquet` resolves to our
  # advertised MIME type (`application/vnd.apache.parquet`) on distros that don't
  # include it by default.
  local parquet_mime_def
  parquet_mime_def="$appdir/usr/share/mime/packages/${EXPECTED_MIME_DEFINITION_BASENAME}"
  if [ ! -f "$parquet_mime_def" ]; then
    die "Missing Parquet shared-mime-info definition in AppImage: squashfs-root/usr/share/mime/packages/${EXPECTED_MIME_DEFINITION_BASENAME}"
  fi
  if ! grep -Fq 'application/vnd.apache.parquet' "$parquet_mime_def" || ! grep -Fq '*.parquet' "$parquet_mime_def"; then
    die "Parquet shared-mime-info definition file is missing expected content in AppImage: squashfs-root/usr/share/mime/packages/${EXPECTED_MIME_DEFINITION_BASENAME}"
  fi

  local applications_dir
  applications_dir="$appdir/usr/share/applications"
  if [ ! -d "$applications_dir" ]; then
    die "Missing applications directory: squashfs-root/usr/share/applications/"
  fi

  declare -a desktop_files=()
  while IFS= read -r -d '' desktop_file; do
    desktop_files+=("$desktop_file")
  done < <(find "$applications_dir" -maxdepth 2 -type f -name '*.desktop' -print0 2>/dev/null || true)

  if [ "${#desktop_files[@]}" -eq 0 ]; then
    die "No .desktop files found under squashfs-root/usr/share/applications/"
  fi

  # Filter to the desktop entry (or entries) that actually launch our app. It's
  # possible (though uncommon) for an AppImage AppDir to contain unrelated
  # .desktop entries; we should validate *our* desktop integration.
  local expected_binary_escaped
  expected_binary_escaped="$(printf '%s' "$EXPECTED_MAIN_BINARY" | sed -e 's/[][(){}.^$*+?|\\]/\\&/g')"
  local exec_token_re
  # Accept either the expected main binary name or AppRun (common in AppImage .desktop entries).
  exec_token_re="(^|[[:space:]])[\"']?([^[:space:]]*/)?(AppRun|${expected_binary_escaped})[\"']?([[:space:]]|$)"

  declare -a matched_desktop_files=()
  local desktop_file
  for desktop_file in "${desktop_files[@]}"; do
    local exec_line
    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    local exec_value
    exec_value="$(printf '%s' "$exec_line" | sed -E "s/^[[:space:]]*Exec[[:space:]]*=[[:space:]]*//I")"
    if [[ -n "$exec_value" ]] && printf '%s' "$exec_value" | grep -Eq "${exec_token_re}"; then
      matched_desktop_files+=("$desktop_file")
    fi
  done

  if [ "${#matched_desktop_files[@]}" -eq 0 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop files appear to target the expected executable '${EXPECTED_MAIN_BINARY}' (or AppRun) in their Exec= entry." >&2
    echo "${SCRIPT_NAME}: error: .desktop files inspected:" >&2
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#$appdir/}"
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [ -z "$exec_line" ]; then
        exec_line="(no Exec= entry)"
      fi
      echo "  - ${rel}: ${exec_line}" >&2
    done
    die "No .desktop files appear to target the expected executable '${EXPECTED_MAIN_BINARY}' (or AppRun) in Exec=."
  fi

  desktop_files=("${matched_desktop_files[@]}")

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
  local has_parquet_mime=0
  local bad_exec_count=0
  local required_parquet_mime="application/vnd.apache.parquet"
  # Track which expected URL scheme handlers appear in the desktop entry/entries.
  declare -A found_scheme_mimes=()

  for desktop_file in "${desktop_files[@]}"; do
    local mime_line
    mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    if [ -z "$mime_line" ]; then
      continue
    fi

    has_any_mimetype=1

    local mime_value
    mime_value="$(printf '%s' "$mime_line" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//")"

    if printf '%s' "$mime_value" | grep -Fqi "$required_parquet_mime"; then
      has_parquet_mime=1
    fi

    # Tokenize MimeType= (semicolon-delimited) so scheme handler checks are exact matches,
    # not substring matches (e.g. avoid treating x-scheme-handler/<scheme>-extra as matching
    # expected x-scheme-handler/<scheme>).
    local -a mime_tokens=()
    IFS=';' read -r -a mime_tokens <<<"$mime_value"
    local -A mime_set=()
    local token
    for token in "${mime_tokens[@]}"; do
      token="${token#"${token%%[![:space:]]*}"}"
      token="${token%"${token##*[![:space:]]}"}"
      token="$(printf '%s' "$token" | tr '[:upper:]' '[:lower:]')"
      if [ -n "$token" ]; then
        mime_set["$token"]=1
      fi
    done

    local has_any_expected_scheme_in_file=0
    local scheme_mime
    for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
      if [ -n "${mime_set["$scheme_mime"]+x}" ]; then
        found_scheme_mimes["$scheme_mime"]=1
        has_any_expected_scheme_in_file=1
      fi
    done

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

    if [ "$has_any_expected_scheme_in_file" -eq 1 ]; then
      # URL scheme handlers also require a placeholder token (%U/%u/%F/%f) in Exec= so the
      # OS passes the opened URL into the app.
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [ -z "$exec_line" ]; then
        echo "${SCRIPT_NAME}: error: ${desktop_file#$appdir/} is missing an Exec= entry (required for URL scheme handlers)" >&2
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        echo "${SCRIPT_NAME}: error: ${desktop_file#$appdir/} Exec= does not include a URL placeholder (%U/%u/%F/%f): ${exec_line}" >&2
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi
  done

  if [ "$has_any_mimetype" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop file contained a MimeType= entry for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: This usually means Linux file associations were not included in the bundle." >&2
    echo "${SCRIPT_NAME}: error: Check ${TAURI_CONF_PATH} → bundle.fileAssociations." >&2
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

  if [ "$has_parquet_mime" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop MimeType= entry advertised Parquet support for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: Expected MimeType= to include '${required_parquet_mime}'." >&2
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
    die "No .desktop file advertised Parquet support for AppImage: $appimage_path"
  fi

  if [ "$has_spreadsheet_mime" -ne 1 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop MimeType= entry advertised spreadsheet/CSV file association support for AppImage: $appimage_path" >&2
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

  declare -a missing_scheme_mimes=()
  for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
    if [ -z "${found_scheme_mimes["$scheme_mime"]+x}" ]; then
      missing_scheme_mimes+=("$scheme_mime")
    fi
  done
  if [ "${#missing_scheme_mimes[@]}" -ne 0 ]; then
    echo "${SCRIPT_NAME}: error: No .desktop MimeType= entry advertised the expected URL scheme handler(s) for AppImage: $appimage_path" >&2
    echo "${SCRIPT_NAME}: error: Missing scheme handler(s): ${missing_scheme_mimes[*]}" >&2
    echo "${SCRIPT_NAME}: error: Expected MimeType= to include: ${EXPECTED_SCHEME_MIMES[*]}" >&2
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
    die "Missing expected URL scheme handler(s) in .desktop file(s) for AppImage: $appimage_path"
  fi

  if [ "$has_xlsx_mime" -ne 1 ]; then
    info "warn: No .desktop file explicitly listed xlsx MIME '${required_xlsx_mime}'. Spreadsheet MIME types were present, but .xlsx double-click integration may be incomplete."
  fi

  # Prefer strict validation against tauri.conf.json so we catch missing MIME types,
  # scheme handlers, compliance artifacts, and Parquet shared-mime-info wiring in the
  # extracted AppImage payload (not just in config).
  #
  # This mirrors the `.deb` / `.rpm` validations in the release workflow.
  if command -v python3 >/dev/null 2>&1; then
    info "Running strict Linux desktop integration verifier against extracted AppImage payload"
    if ! python3 "$REPO_ROOT/scripts/ci/verify_linux_desktop_integration.py" \
      --package-root "$appdir" \
      --tauri-config "$TAURI_CONF_PATH" \
      --expected-main-binary "$EXPECTED_MAIN_BINARY" \
      --doc-package-name "$EXPECTED_MAIN_BINARY"; then
      die "Linux desktop integration verification failed for AppImage: $appimage_path"
    fi
  else
    info "Note: python3 not found; skipping strict desktop integration verification (scripts/ci/verify_linux_desktop_integration.py)"
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
